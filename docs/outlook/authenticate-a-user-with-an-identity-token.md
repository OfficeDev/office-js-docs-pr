---
title: Authenticate Outlook users with Exchange identity tokens
description: Learn how to use an Exchange identity token from an Outlook add-in to identify users and implement single sign-on with your back-end service.
ms.date: 07/10/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Authenticate Outlook users with Exchange identity tokens

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

Use Exchange identity tokens to recognize Outlook add-in users and sign them in to your service without a separate sign-in prompt. The add-in gets a token from Exchange, sends it to your back-end service, and your service validates it before treating the request as authenticated. For guidance on when to use this token type, see [Exchange user identity token](authentication.md#exchange-user-identity-token).

> [!IMPORTANT]
> This is just a simple example of an SSO implementation. As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.

## How token-based sign-in works

1. Your add-in requests an Exchange user identity token.
2. Your add-in sends that token to your back-end service with each request.
3. Your back-end validates the token and maps it to a user in your system.

## Send the ID token with each request

Obtain the Exchange user identity token from the server by calling [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getuseridentitytokenasync-member(1)). Then send this token with every request to your back-end service, such as in a header or in the request body.

## Validate the token

The back-end MUST validate the token before accepting it. This is an important step to ensure that the token was issued by the user's Exchange server. For information on validating Exchange user identity tokens, see [Validate an Exchange identity token](validate-an-identity-token.md).

Once validated and decoded, the payload of the token looks something like the following:

```json
{ 
    "aud" : "https://mailhost.contoso.com/IdentityTest.html",
    "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com",
    "nbf" : "1505749527",
    "exp" : "1505778327",
    "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
    "isbrowserhostedapp":"true",
    "appctx" : {
        "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389",
        "version" : "ExIdTok.V1",
        "amurl" : "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
    }
}
```

## Map the token to a user in your back-end

Your back-end service can calculate a unique user ID from the token and map it to a user in your internal user system. For example, if you store users in a database, you can save this unique ID in each user's record.

### Generate a unique ID

Use a combination of the `msexchuid` and `amurl` properties. For example, concatenate the two values and generate a Base64-encoded string. Because this value is generated consistently from the token, you can map an Exchange user identity token back to the same user in your system.

### Check the user

With the unique ID generated, the next step is to check for a user in your system with that associated ID.

- If the user is found, the back-end treats the request as authenticated, and allows the request to proceed.
- If the user is not found, then the back-end returns an error indicating that the user needs to sign in. The add-in then prompts the user to sign in to the back-end using your existing authentication method. Once the user is authenticated, the Exchange user identity token is submitted with the user authentication details. The back-end can then update the user's record in your system with the unique ID.
