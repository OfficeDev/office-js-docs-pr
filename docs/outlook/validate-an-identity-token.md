---
title: Validate an Outlook add-in identity token
description: Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect.
ms.date: 03/21/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Validate an Exchange identity token

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect. Exchange user identity tokens are JSON Web Tokens (JWT). The steps required to validate a JWT are described in [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).

We suggest that you use a four-step process to validate the identity token and obtain the user's unique identifier. First, extract the JSON Web Token (JWT) from a Base64-encoded URL string. Second, make sure that the token is well-formed, that it is for your Outlook add-in, that it has not expired, and that you can extract a valid URL for the authentication metadata document. Next, retrieve the authentication metadata document from the Exchange server and validate the signature attached to the identity token. Finally, compute a unique identifier for the user by concatenating the user's Exchange ID with the URL of the authentication metadata document.

## Extract the JSON Web Token

The token returned from [getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) is an encoded string representation of the token. In this form, per RFC 7519, all JWTs have three parts, separated by a period. The format is as follows.

```json
{header}.{payload}.{signature}
```

The header and payload should be Base64-decoded to obtain a JSON representation of each part. The signature should be Base64-decoded to obtain a byte array containing the binary signature.

For more information about the contents of the token, see [Inside the Exchange identity token](inside-the-identity-token.md).

After you have the three decoded components, you can proceed with validating the content of the token.

## Validate token contents

To validate the token contents, you should check the following:

- Check the header and verify that the:
  - `typ` claim is set to `JWT`.
  - `alg` claim is set to `RS256`.
  - `x5t` claim is present.

- Check the payload and verify that the:
  - `amurl` claim inside the `appctx` is set to the location of an authorized token signing key manifest file. For example, the expected `amurl` value for Microsoft 365 is https://outlook.office365.com:443/autodiscover/metadata/json/1. See the next section [Verify the domain](#verify-the-domain) for additional information.
  - Current time is between the times specified in the `nbf` and `exp` claims. The `nbf` claim specifies the earliest time that the token is considered valid, and the `exp` claim specifies the expiration time for the token. It is recommended to allow for some variation in clock settings between servers.
  - `aud` claim is the expected URL for your add-in.
  - `version` claim inside the `appctx` claim is set to `ExIdTok.V1`.

### Verify the domain

When implementing the verification logic described in the previous section, you must also require that the domain of the `amurl` claim matches the Autodiscover domain for the user. To do so, you'll need to use or implement [Autodiscover for Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).

- For Exchange Online, confirm that the `amurl` is a well-known domain (https://outlook.office365.com:443/autodiscover/metadata/json/1), or belongs to a geo-specific or specialty cloud ([Office 365 URLs and IP address ranges](/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide&preserve-view=true)).

- If your add-in service has a preexisting configuration with the user's tenant, then you can establish if this `amurl` is trusted.

- For an [Exchange hybrid deployment](/microsoft-365/enterprise/configure-exchange-server-for-hybrid-modern-authentication?view=o365-worldwide&preserve-view=true), use OAuth-based Autodiscover to verify the domain expected for the user. However, while the user will need to authenticate as part of the Autodiscover flow, your add-in should never collect the user's credentials and do basic authentication.

If your add-in can't verify the `amurl` using any of these options, you may choose to have your add-in shut down gracefully with an appropriate notification to the user if authentication is necessary for the add-in's workflow.

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

The available signing keys are in the `keys` array. Select the correct key by ensuring that the `x5t` value in the `keyinfo` property matches the `x5t` value in the header of the token. The public key is inside the `value` property in the `keyvalue` property, stored as a Base64-encoded byte array.

After you have the correct public key, verify the signature. The signed data is the first two parts of the encoded token, separated by a period:

```json
{header}.{payload}
```

## Compute the unique ID for an Exchange account

Create a unique identifier for an Exchange account by concatenating the authentication metadata document URL with the Exchange identifier for the account. When you have this unique identifier, use it to create a single sign-on (SSO) system for your Outlook add-in web service. For details about using the unique identifier for SSO, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).

## Use a library to validate the token

There are a number of libraries that can do general JWT parsing and validation. Microsoft provides the `System.IdentityModel.Tokens.Jwt` library that can be used to validate Exchange user identity tokens.

> [!IMPORTANT]
> We no longer recommend the Exchange Web Services Managed API because the Microsoft.Exchange.WebServices.Auth.dll, though still available, is now obsolete and relies on unsupported libraries like Microsoft.IdentityModel.Extensions.dll.

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

## See also

- [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
