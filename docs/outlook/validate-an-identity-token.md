
# Validate an Exchange identity token

Your Outlook add-in can send you an identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect. The examples in this article show you how to validate the Exchange identity token using a validation object written in C#; however, you can use any programming language to do the validation. The steps required to validate the token are described in the [JSON Web Token (JWT) Internet Draft](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl). 

We suggest that you use a four-step process to validate the identity token and obtain the user's unique identifier. First, extract the JSON Web Token (JWT) from a base64 URL-encoded string. Second, make sure that the token is well-formed, that it is for your Outlook add-in, that it has not expired, and that you can extract a valid URL for the authentication metadata document. Next, retrieve the authentication metadata document from the Exchange server and validate the signature attached to the identity token. Finally, compute a unique identifier for the user by hashing the user's Exchange ID with the URL of the authentication metadata document. Overall the process may seem complex, but each individual step is quite simple.
You can download the solution that contains these examples from the web at  [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken).
 




## Set up to validate your identity token


The code examples in this article depend on the Windows Identity Foundation (WIF), along with a DLL that extends the WIF with handlers for JSON tokens. You can download the required assemblies from the following locations:


- [Windows Identity Foundation](http://msdn.microsoft.com/en-us/security/aa570351)
    
- [Windows.IdentityModel.Extensions.dll for 32-bit applications](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [Windows.IdentityModel.Extensions.dll for 64-bit applications](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## Extract the JSON Web Token


The  **Decode** factory method splits the JWT from the Exchange server into the three strings that make up the token, and then uses the **Base64Decode** method (shown in the second example) to decode the JWT header and payload into JSON strings. The strings are passed to the **JsonToken** constructor, where the contents of the JWT are validated and a new **JsonToken** object instance is returned.


```C#
    public static JsonToken Decode(string rawToken)
    {
      string[] tokenParts = rawToken.Split('.');

      if (tokenParts.Length != 3)
      {
        throw new ApplicationException("Token must have three parts separated by '.' characters.");
      }

      string encodedHeader = tokenParts[0];
      string encodedPayload = tokenParts[1];
      string signature = tokenParts[2];

      string decodedHeader = Base64Decode(encodedHeader);
      string decodedPayload = Base64Decode(encodedPayload);

      JavaScriptSerializer serializer = new JavaScriptSerializer();

      Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
      Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

      return new JsonToken(header, payload, signature);
    }
```

The  **Base64Decode** method implements the decoding logic that is described in the "Notes on implementing base64url encoding without padding" appendix in the [JSON Web Token (JWT) Internet Draft](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl).




```C#
    public static Encoding TextEncoding = Encoding.UTF8;

    private static char Base64PadCharacter = '=';
    private static char Base64Character62 = '+';
    private static char Base64Character63 = '/';
    private static char Base64UrlCharacter62 = '-';
    private static char Base64UrlCharacter63 = '_';

    private static byte[] DecodeBytes(string arg)
    {
      if (String.IsNullOrEmpty(arg))
      {
        throw new ApplicationException("String to decode cannot be null or empty.");
      }

      StringBuilder s = new StringBuilder(arg);
      s.Replace(Base64UrlCharacter62, Base64Character62);
      s.Replace(Base64UrlCharacter63, Base64Character63);

      int pad = s.Length % 4;
      s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

      return Convert.FromBase64String(s.ToString());
    }

    private static string Base64Decode(string arg)
    {
      return TextEncoding.GetString(DecodeBytes(arg));
    }
```


## Parse the JWT


The constructor for the  **JsonToken** object checks the structure and contents of the JWT to determine whether it's valid. It's best to do this before you ask for the authentication metadata document. If the JWT does not contain the proper claims, or if it's outside of its lifetime, you can avoid a call to the Exchange server and the associated delay.

The constructor calls utility methods to determine whether the different claims are present and in scope. If there is a problem, the utility method will throw an application exception. If no exceptions are thrown, the  **IsValid** property is set to **true** and the token is ready for signature validation.

Each of the utility methods is described further later in this article.




```C#
    public JsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
    {

      // Assume that the token is invalid to start out.
      this.IsValid = false;

      // Set the private dictionaries that contain the claims.
      this.headerClaims = header;
      this.payloadClaims = payload;
      this.signature = signature;

      // If there is no "appctx" claim in the token, throw an ApplicationException.
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.AppContext))
      {
        throw new ApplicationException(String.Format("The {0} claim is not present.", AuthClaimTypes.AppContext));
      }

      appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload[AuthClaimTypes.AppContext]);


      // Validate the header fields.
      this.ValidateHeader();

      // Determine whether the token is within its valid time.
      this.ValidateLifetime();

      // Validate that the token was sent to the correct URL.
      this.ValidateAudience();

      // Validate the token version.
      this.ValidateVersion();

      // Make sure that the appctx contains an authentication
      // metadata location.
      this.ValidateMetadataLocation();

      // If the token passes all the validation checks, we
      // can assume that it is valid.
      this.IsValid = true;
    }
```


### ValidateHeader method

The  **ValidateHeader** method checks to make sure that the required claims are in the token header, and that the claims have the correct values. The header must be set as follows; otherwise, the method will throw an application exception and end.

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "<thumbprint>" }
```

```C#
    private void ValidateHeaderClaim(string key, string value)
    {
      if (!this.headerClaims.ContainsKey(key))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
      }

      if (!value.Equals(this.headerClaims[key]))
      {
        throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
      }
    }

    private void ValidateHeader()
    {
      ValidateHeaderClaim(AuthClaimTypes.TokenType, Config.TokenType);
      ValidateHeaderClaim(AuthClaimTypes.Algorithm, Config.Algorithm);
    
      if (!this.headerClaims.ContainsKey(AuthClaimTypes.x509Thumprint))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", AuthClaimTypes.x509Thumprint));
      }
    }


```


### ValidateLifetime method

Two dates are provided in the JWT: "nbf" ("not before") gives the date and time that the token becomes valid, and "exp" gives the time that the token expires. Only tokens presented between these two dates should be considered valid. To accommodate minor differences in the clock setting between the server and the client, this method will validate tokens up to five minutes before and 5 minutes after the times in the token.


```C#
    private void ValidateLifetime()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidFrom))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidFrom));
      }

      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidTo))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidTo));
      }

      DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0,DateTimeKind.Utc);

      TimeSpan padding = new TimeSpan(0, 5, 0);

      DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidFrom]));
      DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidTo]));

      DateTime now = DateTime.UtcNow;

      if (now < (validFrom - padding))
      {
        throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
      }

      if (now > (validTo + padding))
      {
        throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
      }
    }
```

The  **validFrom** ("nbf") and **validTo** ("exp") dates are sent as the number of seconds since the Unix epoch, January 1, 1970. The dates and times are calculated using UTC to avoid any problems with time zone differences between the Exchange server and the server running the validation code.


### ValidateAudience method

The identity token is only valid for the add-in that requested it. The  **ValidateAudience** method checks the audience claim in the token to ensure that it matches the expected URL for the Outlook add-in.


```C#
    private void ValidateAudience()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.Audience))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", AuthClaimTypes.Audience));
      }

      string location = Config.Audience.Replace("/", "-").Replace("\\", "-");
      string audience = this.payloadClaims[AuthClaimTypes.Audience].Replace("/", "-").Replace("\\", "-");

      if (!location.Equals(audience))
      {
        throw new ApplicationException(String.Format(
          "The audience URL does not match. Expected {0}; got {1}.",
          Config.Audience, this.payloadClaims[AuthClaimTypes.Audience]));
      }
    }

```


### ValidateVersion method

The  **ValidateVersion** method checks the version of the identity token and makes sure that it matches the expected version. Different versions of the token can carry different claims. Checking the version ensures that the expected claims will be in the identity token.


```js
    private void ValidateVersion()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchExtensionVersion))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchExtensionVersion));
      }

      if (!Config.Version.Equals(this.appContext[AuthClaimTypes.MsExchExtensionVersion]))
      {
        throw new ApplicationException(String.Format(
          "The version does not match. Expected {0}; got {1}.",
          Config.Version, this.appContext[AuthClaimTypes.MsExchExtensionVersion]));
      }
    }

```


### ValidateMetadataLocation method

The authentication metadata object that is stored on the Exchange server contains the information that is required to validate the signature included in the identity token. The  **ValidateMetadataLocation** method makes sure that there is an authentication metadata URL claim in the identity token, actually validating the signature takes place in the next step.


```C#
    private void ValidateMetadataLocation()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchAuthMetadataUrl))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchAuthMetadataUrl));
      }
    }

```


## Validate the identity token signature


After you know that the JWT contains the claims that you need to validate the signature, you can use the Windows Identity Foundation (WIF) and the WIF extensions to validate the signature on the token. You need the following information to validate the signature:


- The original base-64 URL-encoded identity token string sent from the Exchange server.
    
- The authentication metadata document location from the JWT.
    
- The audience URL from the JWT.
    
In this example, the constructor for an  **IdentityToken** object gets the authentication metadata document from the Exchange server and validates the signature on the identity token. If the identity token is valid, you can use the **IdentityToken** object instance to get the unique user ID that is included in the identity token.




```C#
    public IdentityToken(string rawToken, string audience, string authMetadataEndpoint)
    {
      X509Certificate2 currentCertificate = null;

      currentCertificate = AuthMetadata.GetSigningCertificate(new Uri(authMetadataEndpoint));

      JsonWebSecurityTokenHandler jsonTokenHandler =
          GetSecurityTokenHandler(audience, authMetadataEndpoint, currentCertificate);

      SecurityToken jsonToken = jsonTokenHandler.ReadToken(rawToken);
      JsonWebSecurityToken webToken = (JsonWebSecurityToken)jsonToken;

      SigningCertificateThumbprint = currentCertificate.Thumbprint;
      Issuer = webToken.Issuer;
      Audience = webToken.Audience;
      ValidTo = webToken.ValidTo;
      ValidFrom = webToken.ValidFrom;
      foreach (JsonWebTokenClaim claim in webToken.Claims)
      {
        if (claim.ClaimType.Equals(AuthClaimTypes.AppContextSender))
        {
          ApplicationContextSender = claim.Value;
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.IsBrowserHostedApp))
        {
          IsBrowserHostedApp = claim.Value == "true";
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.AppContext))
        {
          string[] appContextClaims = claim.Value.Split(',');
          Dictionary<string, string> appContext =
              new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(claim.Value);
          AuthenticationMetaDataUrl = appContext[AuthClaimTypes.MsExchAuthMetadataUrl];
          ExchangeID = appContext[AuthClaimTypes.MsExchImmutableId];
          TokenVersion = appContext[AuthClaimTypes.MsExchTokenVersion];
        }
      }
    }


```

Most of the code in the  **IdentityToken** object constructor sets the properties on the instance with the claims from the Exchange server. The constructor calls the **GetSecurityTokenHandler** method to get a token handler that will validate the Exchange identity token. The **GetSecurityTokenHandler** method calls two utility methods, **GetMetadataDocument** and **GetSigningCertificate**, which do the work of getting the signing certificate from the Exchange server. Each of these methods is described in the following sections.


### GetSecurityTokenHandler method

The  **GetSecurityTokenHandler** method returns a WIF token handler that will validate the identity token. Most of the code in the method initializes the token handler to do the validation; however, the method does call the **GetSigningCertificate** method to retrieve the X.509 certificate used to sign the token from the Exchange server.


```C#
    private JsonWebSecurityTokenHandler GetSecurityTokenHandler(string audience,
        string authMetadataEndpoint,
        X509Certificate2 currentCertificate)
    {
      JsonWebSecurityTokenHandler jsonTokenHandler = new JsonWebSecurityTokenHandler();
      jsonTokenHandler.Configuration = new SecurityTokenHandlerConfiguration();

      jsonTokenHandler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Always);
      jsonTokenHandler.Configuration.AudienceRestriction.AllowedAudienceUris.Add(
        new Uri(audience, UriKind.RelativeOrAbsolute));

      jsonTokenHandler.Configuration.CertificateValidator = X509CertificateValidator.None;

      jsonTokenHandler.Configuration.IssuerTokenResolver =
        SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
          new ReadOnlyCollection<SecurityToken>(new List<SecurityToken>(
            new SecurityToken[]
            {
              new X509SecurityToken(currentCertificate)
            })), false);

      ConfigurationBasedIssuerNameRegistry issuerNameRegistry = new ConfigurationBasedIssuerNameRegistry();
      issuerNameRegistry.AddTrustedIssuer(currentCertificate.Thumbprint, Config.ExchangeApplicationIdentifier);
      jsonTokenHandler.Configuration.IssuerNameRegistry = issuerNameRegistry;

      return jsonTokenHandler;
    }
```


### GetSigningCertificate method

The  **GetSigningCertificate** method calls the **GetMetadataDocument** method to retrieve the authentication metadata from the Exchange server and then returns the first X.509 certificate in the authentication metadata document. If the document doesn't exist, the method throws an application exception.


```C#
    private X509Certificate2 GetSigningCertificate(Uri authMetadataEndpoint)
    {
      JsonAuthMetadataDocument document = GetMetadataDocument(authMetadataEndpoint);

      if (null != document.keys &amp;&amp; document.keys.Length > 0)
      {
        JsonKey signingKey = document.keys[0];

        if (null != signingKey &amp;&amp; null != signingKey.keyValue)
        {
          return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
        }
      }

      throw new ApplicationException("The metadata document does not contain a signing certificate.");
    }

```


### GetMetadataDocument method

The authentication metadata document contains the information that you need to validate the signature on the Exchange identity token. The document is sent as a JSON string. The  **GetMetatDataDocument** method requests the document from the location specified in the Exchange identity token and returns an object that encapsulates the JSON string as an object. If the URL does not contain an authentication metadata document, the method throws an application exception.


```C#
    private JsonAuthMetadataDocument GetMetadataDocument(Uri authMetadataEndpoint)
    {
      // Uncomment the next line if your Exchange server uses the default
      // self-signed certificate.
      // ServicePointManager.ServerCertificateValidationCallback = Config.CertificateValidationCallback;

      byte[] acsMetadata;
      using (WebClient webClient = new WebClient())
      {
        acsMetadata = webClient.DownloadData(authMetadataEndpoint);
      }
      string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

      JsonAuthMetadataDocument document = new JavaScriptSerializer().Deserialize<JsonAuthMetadataDocument>(jsonResponseString);

      if (null == document)
      {
        throw new ApplicationException(String.Format("No authentication metadata document found at {0}.", authMetadataEndpoint));
      }

      return document;
    }
```

By default the Exchange server uses a self-signed X.509 certificate to authenticate requests for the authentication metadata document. Unless you install a certificate that traces back to a root server, you must create a certificate validation callback method or else the request for the authentication metadata document will fail. 

The  **ServicePointManager** class in the .NET Framework System.Net namespace enables you to hook up a validation callback method by setting the **ServerCertificateValidationCallback** property. You can see an example of a certificate validation callback method that is suitable for development and testing in the article [Validating X509 certificates](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx).


 **Security Note**  If you use a certificate validation callback method, you must make sure that it meets the security requirements of your organization.


## Compute the unique ID for an Exchange account


You can create a unique identifier for an Exchange account by hashing the authentication metadata document URL with the Exchange identifier for the account. When you have this unique identifier, you can use it to create a single sign-on (SSO) system for your Outlook add-in web service. For details about using the unique identifier for SSO, see [Authenticate a user with an identity token for Exchange](../outlook/authenticate-a-user-with-an-identity-token.md)

The  **UniqueUserIdentification** property creates a salted SHA256 hash of the Exchange ID and authentication metadata URL by using the standard SHA256 provider from the **System.Security.Cryptography** namespace.


 **Security Note**  You must hash the authentication metadata document with the Exchange ID to create a unique identifier for an account. Using just the Exchange ID can expose your service to unauthorized users. And as always when dealing with authentication and security, you must make sure that using the unique identifier created with this method meets the security requirements of your application.




```C#
    // Salt to apply when creating unique ID.
    private byte[] Salt = new byte[] {<Provide random salt bytes here };

    private string ComputeUniqueIdentification()
    {
      byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(ExchangeID, AuthenticationMetaDataUrl));

      // Combine input bytes and salt.
      byte[] saltedInput = new byte[Salt.Length + inputBytes.Length];
      Salt.CopyTo(saltedInput, 0);
      inputBytes.CopyTo(saltedInput, Salt.Length);

      // Compute the unique key.
      byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

      // Convert the hashed value to a string and return.
      return BitConverter.ToString(hashedBytes);
    }

    public string UniqueUserIdentification
    {
      get { return ComputeUniqueIdentification(); }
    }


```


## Utility objects


The code examples in this article depend on a few utility objects that provide friendly names to the constants that are used. The following table lists the utility objects.


**Table 1: Utility objects**


|**Object**|**Description**|
|:-----|:-----|
|**AuthClaimsType**|Collects the claim identifiers that are used by the token validation code into a single place.|
|**Config**|Provides the constants to validate the identity token. |
|**JsonAuthMetadataDocument**|Encapsulates the JSON authentication metadata document sent from the Exchange server.|

### AuthClaimTypes object

The  **AuthClaimTypes** object collects the claim identifiers that are used by the token validation code into a single place. It includes both standard JWT claims as well as the specific claims in the Exchange identity token.


```C#
  public class AuthClaimTypes
  {
    public const string NameIdentifier =
        JsonWebTokenConstants.ReservedClaims.NameIdentifier;
    public const string MsExchImmutableId = "msexchuid";
    public const string MsExchTokenVersion = "version";
    public const string MsExchAuthMetadataUrl = "amurl";

    public const string AppContext =
        JsonWebTokenConstants.ReservedClaims.AppContext;
    public const string Audience =
        JsonWebTokenConstants.ReservedClaims.Audience;
    public const string Issuer =
        JsonWebTokenConstants.ReservedClaims.Issuer;
    public const string ValidFrom =
        JsonWebTokenConstants.ReservedClaims.NotBefore;
    public const string ValidTo =
        JsonWebTokenConstants.ReservedClaims.ExpiresOn;

    public const string AppContextSender = "appctxsender";
    public const string IsBrowserHostedApp = "isbrowserhostedapp";

    public const string TokenType = "typ";
    public const string Algorithm = "alg";
    public const string x509Thumbprint = "x5t";      
  }
```


### Config object

The  **Config** object contains the constants that are used to validate the identity token, as well as a certificate validation callback method that you can use if your server does not have an X509 certificate that traces back to a root certificate.


 **Security Note**  The security certificate callback method is only required if your server uses the default self-signed certificate. The callback method in this example returns  **false** when the certificate is self-signed, so you'll need to replace it with a callback method that meets the security requirements of your organization. For an example of a certificate validation callback method that is suitable for development and testing, see [Validating X509 certificates](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx).


```C#
  public static class Config
  {
    public static string Algorithm = "RS256";
    public static string Audience = @"https:\\localhost:44300\Pages\IdentityTest.html";
    public static string TokenType = "JWT";
    public static string Version = "ExIdTok.V1";

    public static string ExchangeApplicationIdentifier = "Exchange";

    internal static bool CertificateValidationCallback(
    object sender,
    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
    System.Security.Cryptography.X509Certificates.X509Chain chain,
    System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      else
      {
        return false;
      }
    }
  }
```


### JsonAuthMetadataDocument object

The  **JsonAuthMetadataDocument** object exposes the contents of the authentication metadata document through properties.


```C#
using System;

namespace IdentityTest
{
  public class JsonAuthMetadataDocument
  {
    public string id { get; set; }
    public string version { get; set; }
    public string name { get; set; }
    public string realm { get; set; }
    public string serviceName { get; set; }
    public string issuer { get; set; }
    public string [] allowedAudiences { get; set; }
    public JsonKey[] keys;
    public JsonEndpoint[] endpoints;
  }

  public class JsonEndpoint
  {
    public string location { get; set; }
    public string protocol { get; set; }
    public string usage { get; set; }
  }

  public class JsonKey
  {
    public string usage { get; set; }
    public JsonKeyValue keyValue { get; set; }
  }

  public class JsonKeyValue
  {
    public string type { get; set; }
    public string value { get; set; }
  }
}

```


## Additional resources



- [Authenticate an Outlook add-in by using Exchange identity tokens](../outlook/authentication.md)
    
- [Inside the Exchange identity token](../outlook/inside-the-identity-token.md)
    
