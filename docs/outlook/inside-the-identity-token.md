---
title: Inside the Exchange identity token in an Outlook add-in
description: Learn about the contents of an Exchange user identity token generated from an Outlook add-in.
ms.date: 04/12/2024
ms.localizationpriority: medium
---

# Inside the Exchange identity token

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

The Exchange user identity token returned by the [getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method provides a way for your add-in code to include the user's identity with calls to your back-end service. This article will discuss the format and contents of the token.

An Exchange user identity token is a base-64 URL-encoded string that is signed by the Exchange server that sent it. The token is not encrypted, and the public key that you use to validate the signature is stored on the Exchange server that issued the token. The token has three parts: a header, a payload, and a signature. In the token string, the parts are separated by a period character (`.`) to make it easy for you to split the token.

Exchange uses a the JSON Web Token (JWT) format for the identity token. For information about JWT tokens, see [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).

## Identity token header

The header provides information about the format and signature information of the token. The following example shows what the header of the token looks like.

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
The following table describes the parts of the token header.

| Claim | Value | Description |
|:-----|:-----|:-----|
| `typ` | `JWT` | Identifies the token as a JSON Web Token. All identity tokens provided by Exchange server are JWT tokens. |
| `alg` | `RS256` | The hashing algorithm that is used to create the signature. All tokens provided by Exchange server use the RSASSA-PKCS1-v1_5 with SHA-256 hash algorithm. |
| `x5t` | Certificate thumbprint | The X.509 thumbprint of the token. |

## Identity token payload

The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.

```JSON
{ 
  "aud": "https://mailhost.contoso.com/IdentityTest.html", 
  "iss": "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
  "nbf": "1331579055", 
  "exp": "1331607855", 
  "appctxsender": "00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
  "isbrowserhostedapp": "true",
  "appctx": { 
    "msexchuid": "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com",
    "version": "ExIdTok.V1",
    "amurl": "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
  } 
}
```

<br/>
 
The following table lists the parts of the identity token payload.

| Claim | Description |
|:-----|:-----|
| `aud` | The URL of the add-in that requested the token. A token is only valid if it is sent from the add-in that is running in the client's webview control. The URL of the add-in is specified in the manifest. The markup depends on the type of manifest.</br></br>**Add-in only manifest:** If the add-in uses the Office Add-ins manifests schema v1.1, this URL is the URL specified in the first `<SourceLocation>` element, under the form type `ItemRead` or `ItemEdit`, whichever occurs first as part of the [FormSettings](/javascript/api/manifest/formsettings) element in the add-in manifest.</br></br>**Unified manifest for Microsoft 365:** The URL is specified in the `"extensions.audienceClaimUrl"` property. |
| `iss` | A unique identifier for the Exchange server that issued the token. All tokens issued by this Exchange server will have the same identifier. |
| `nbf` | The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970. |
| `exp` | The date and time that the token is valid until. The value is the number of seconds since January 1, 1970. |
| `appctxsender` | A unique identifier for the Exchange server that sent the application context. |
| `isbrowserhostedapp` | Indicates whether the add-in is hosted in a browser. |
| `appctx` | The application context for the token. |

The information in the appctx claim provides you with the unique identifier for the account and the location of the public key used to sign the token. The following table lists the parts of the `appctx` claim.

| Application context property | Description |
|:-----|:-----|
| `msexchuid` | A unique identifier associated with the email account and the Exchange server. |
| `version` | The version number of the token. For all tokens provided by Exchange, the value is `ExIdTok.V1`. |
| `amurl` | The URL of the authentication metadata document that contains the public key of the X.509 certificate that was used to sign the token.<br/><br/>For more information about how to use the authentication metadata document, see [Validate an Exchange identity token](validate-an-identity-token.md). |

## Identity token signature

The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.

## See also

For an example that parses the Exchange user identity token, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).
