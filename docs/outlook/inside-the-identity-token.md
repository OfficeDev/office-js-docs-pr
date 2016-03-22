
# Inside the Exchange identity token
Find out what's inside an Exchange 2013 identity token.



The authentication identity token that the Exchange server sends to your Outlook add-in is opaque to your add-in; you don't have to look inside the token to send it on to your server. But when you're writing the web service code that interacts with your Outlook add-in, you'll need to know what's inside the identity token.

## What is an identity token?


An identity token is a base-64 URL encoded string that is self-signed by the Exchange server that sent it. The token is not encrypted, and the public key that you use to validate the signature is stored on the Exchange server that issued the token. The token has three parts: a header, a payload, and a signature. In the token string, the parts are separated by a "." character to make it easy for you to split the token.

Exchange 2013 uses a JSON Web Token (JWT) for the identity token. For information about JWT tokens, see the [JSON Web Token (JWT) Internet Draft](http://self-issued.info/docs/draft-goland-json-web-token-00.html).


### Identity token header

The header identifies the token and lets your web service know what kind of token is being presented. The following example shows what he header of the token looks like.

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "Un6V7lYN-rMgaCoFSTO5z707X-4" }
```

The following table describes the parts of the identity token header.


**Parts of the identity token header**


|**Claim**|**Value**|**Description**|
|:-----|:-----|:-----|
|typ|"JWT"|Identifies the token as a JSON Web Token. All identity tokens provided by the Exchange server are JWT tokens.|
|alg|"RS256"|The hashing algorithm that is used to create the signature. All tokens provided by the Exchange server use the RS-256 algorithm.|
|x5t|Certificate thumbprint|The X.509 thumbprint of the token.|

### Identity token payload

The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.
```js

{ 
   "aud" : "https://mailhost.contoso.com/IdentityTest.html", 
   "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
   "nbf" : "1331579055", 
   "exp" : "1331607855", 
   "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
   "isbrowserhostedapp":"true",
"appctx" : { 
     "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com" "version" : "ExIdTok.V1" "amurl" :         "https://mailhost.contoso.com:443/autodiscover/metadata/json/1" 
     } 
}
```
The following table lists the parts of the identity token payload.


**Parts of the identity token payload**


|**Claim**|**Description**|
|:-----|:-----|
|aud|The URL of the add-in that requested the token. A token is only valid if it is sent from the add-in that is running in the client's browser. If the add-in uses the Office Add-ins manifests schema v1.1, this URL is the URL specified in the first  **SourceLocation** element, under the form type **ItemRead** or **ItemEdit**, whichever occurs first as part of the [FormSettings](http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx) element in the add-in manifest.|
|iss|A unique identifier for the Exchange server that issued the token. All tokens issued by this Exchange server will have the same identifier.|
|nbf|The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970. |
|exp|The date and time that the token is valid until. The value is the number of seconds since January 1, 1970.|
|appctxsender|A unique identifier for the Exchange server that sent the application context.|
|isbrowserhostedapp|Indicates whether the add-in is hosted in a browser.|
|appctx|The application context for the token. |
The information in the appctx claim provides you with the address of the email account, and a unique identifier for the account. The following table lists the parts of the appctx claim.



|**appctx claim part**|**Description**|
|:-----|:-----|
|msexchuid|A unique identifier associated with the email account and the Exchange server.|
|version|The version number of the token. For all tokens provided by a server that is running Exchange 2013, the value is "ExIdTok.V1".|
|amurl|The URL of the authentication metadata document that contains the public key of the X.509 certificate that was used to sign the token. For more information about how to use the authentication metadata document, see [Validate an Exchange identity token](../outlook/validate-an-identity-token.md).|

### Identity token signature

The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.


## Additional resources



- [Authenticate an Outlook add-in by using Exchange identity tokens](../outlook/authentication.md)
    
- [Call a service from an Outlook add-in by using an identity token in Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Use the Exchange token validation library](../outlook/use-the-token-validation-library.md)
    
- [Validate an Exchange identity token](../outlook/validate-an-identity-token.md)
    
