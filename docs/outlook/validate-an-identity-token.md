---
title: Validate an Outlook add-in identity token
description: Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect.
ms.date: 11/07/2019
localization_priority: Normal
---

# Validate an Exchange identity token

Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect. Exchange user identity tokens are JSON Web Tokens (JWT). The steps required to validate a JWT are described in [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).

We suggest that you use a four-step process to validate the identity token and obtain the user's unique identifier. First, extract the JSON Web Token (JWT) from a base64 URL-encoded string. Second, make sure that the token is well-formed, that it is for your Outlook add-in, that it has not expired, and that you can extract a valid URL for the authentication metadata document. Next, retrieve the authentication metadata document from the Exchange server and validate the signature attached to the identity token. Finally, compute a unique identifier for the user by concatenating the user's Exchange ID with the URL of the authentication metadata document.

## Extract the JSON Web Token

The token returned from [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) is an encoded string representation of the token. In this form, per RFC 7519, all JWTs have three parts, separated by a period. The format is as follows.

```json
{header}.{payload}.{signature}
```

The header and payload should be base64-decoded to obtain a JSON representation of each part. The signature should be base64-decoded to obtain a byte array containing the binary signature.

For more information about the contents of the token, see [Inside the Exchange identity token](inside-the-identity-token.md).

After you have the three decoded components, you can proceed with validating the content of the token.

## Validate token contents

To validate the token contents, you should check the following.

- Check the header and verify that the:
    - `typ` claim is set to `JWT`.
    - `alg` claim is set to `RS256`.
    - `x5t` claim is present.

- Check the payload and verify that the:
    - `amurl` claim inside the `appctx` is set to the location of an authorized token signing key manifest file. For example, the expected `amurl` value for Office 365 is https://outlook.office365.com:443/autodiscover/metadata/json/1. See the next section [Verify the domain](#verify-the-domain) for additional information.
    - Current time is between the times specified in the `nbf` and `exp` claims. The `nbf` claim specifies the earliest time that the token is considered valid, and the `exp` claim specifies the expiration time for the token. It is recommended to allow for some variation in clock settings between servers.
    - `aud` claim is the expected URL for your add-in.
    - `version` claim inside the `appctx` claim is set to `ExIdTok.V1`.

### Verify the domain

When implementing the verification logic described previously in this section, you should also require that the domain of the `amurl` claim matches the Autodiscover domain for the user. To do so, you'll need to use or implement Autodiscover. To learn more, you can start with [Autodiscover for Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).

## Validate the identity token signature

After you know that the JWT contains the required claims, you can proceed with validating the token signature.

### Retrieve the public signing key

The first step is to retrieve the public key that corresponds to the certificate that the Exchange server used to sign the token. The key is found in the authentication metadata document. This document is a JSON file hosted at the URL specified in the `amurl` claim.

The authentication metadata document uses the following format.

```json
{
    "id": "_70b34511-d105-4e2b-9675-39f53305bb01",
    "version": "1.0",
    "name": "Exchange",
    "realm": "*",
    "serviceName": "00000002-0000-0ff1-ce00-000000000000",
    "issuer": "00000002-0000-0ff1-ce00-000000000000@*",
    "allowedAudiences": [
        "00000002-0000-0ff1-ce00-000000000000@*"
    ],
    "keys": [
        {
            "usage": "signing",
            "keyinfo": {
                "x5t": "enh9BJrVPU5ijV1qjZjV-fL2bco"
            },
            "keyvalue": {
                "type": "x509Certificate",
                "value": "MIIHNTCC..."
            }
        }
    ],
    "endpoints": [
        {
            "location": "https://by2pr06mb2229.namprd06.prod.outlook.com:444/autodiscover/metadata/json/1",
            "protocol": "OAuth2",
            "usage": "metadata"
        }
    ]
}
```

The available signing keys are in the `keys` array. Select the correct key by ensuring that the `x5t` value in the `keyinfo` property matches the `x5t` value in the header of the token. The public key is inside the `value` property in the `keyvalue` property, stored as a base64-encoded byte array.

After you have the correct public key, verify the signature. The signed data is the first two parts of the encoded token, separated by a period:

```json
{header}.{payload}
```

## Compute the unique ID for an Exchange account

You can create a unique identifier for an Exchange account by concatenating the authentication metadata document URL with the Exchange identifier for the account. When you have this unique identifier, you can use it to create a single sign-on (SSO) system for your Outlook add-in web service. For details about using the unique identifier for SSO, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).

## Use a library to validate the token

There are a number of libraries that can do general JWT parsing and validation. Microsoft provides two libraries that can be used to validate Exchange user identity tokens.

### System.IdentityModel.Tokens.Jwt

The [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) library can parse the token and also perform the validation, though you will need to parse the `appctx` claim yourself and retrieve the public signing key.

```cs
// Load the encoded token
string encodedToken = "...";
JwtSecurityToken jwt = new JwtSecurityToken(encodedToken);

// Parse the appctx claim to get the auth metadata url
string authMetadataUrl = string.Empty;
var appctx = jwt.Claims.FirstOrDefault(claim => claim.Type == "appctx");
if (appctx != null)
{
    var AppContext = JsonConvert.DeserializeObject<ExchangeAppContext>(appctx.Value);

    // Token version check
    if (string.Compare(AppContext.Version, "ExIdTok.V1", StringComparison.InvariantCulture) != 0) {
        // Fail validation
    }

    authMetadataUrl = AppContext.MetadataUrl;
}

// Use System.IdentityModel.Tokens.Jwt library to validate standard parts
JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
TokenValidationParameters tvp = new TokenValidationParameters();

tvp.ValidateIssuer = false;
tvp.ValidateAudience = true;
tvp.ValidAudience = "{URL to add-in}";
tvp.ValidateIssuerSigningKey = true;
// GetSigningKeys downloads the auth metadata doc and
// returns a List<SecurityKey>
tvp.IssuerSigningKeys = GetSigningKeys(authMetadataUrl);
tvp.ValidateLifetime = true;

try
{
    var claimsPrincipal = tokenHandler.ValidateToken(encodedToken, tvp, out SecurityToken validatedToken);

    // If no exception, all standard checks passed
}
catch (SecurityTokenValidationException ex)
{
    // Validation failed
}
```

<br/>

The `ExchangeAppContext` class is defined as follows:

```cs
using Newtonsoft.Json;

/// <summary>
/// Representation of the appctx claim in an Exchange user identity token.
/// </summary>
public class ExchangeAppContext
{
    /// <summary>
    /// The Exchange identifier for the user
    /// </summary>
    [JsonProperty("msexchuid")]
    public string ExchangeUid { get; set; }

    /// <summary>
    /// The token version
    /// </summary>
    public string Version { get; set; }

    /// <summary>
    /// The URL to download authentication metadata
    /// </summary>
    [JsonProperty("amurl")]
    public string MetadataUrl { get; set; }
}
```

For an example that uses this library to validate Exchange tokens and has an implementation of `GetSigningKeys`, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).

### Microsoft.Exchange.WebServices

The [Exchange Web Services Managed API](https://www.nuget.org/packages/Microsoft.Exchange.WebServices/) can also validate Exchange user identity tokens. Because it is Exchange-specific, it implements all of the necessary logic to parse the `appctx` claim and verify the token version.

```cs
using Microsoft.Exchange.WebServices.Auth.Validation;

AppIdentityToken ValidateIdentityToken(string rawToken, string expectedAudience)
{
    try
    {
        AppIdentityToken appIdToken = AuthToken.Parse(rawToken) as AppIdentityToken;
        appIdToken.Validate(new Uri(expectedAudience));

        // No exception, validation succeeded
        return appIdToken;
    }
    catch (TokenValidationException ex)
    {
        throw new Exception(string.Format("Token validation failed: {0}", ex.Message));
    }
}
```

## See also

- [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
