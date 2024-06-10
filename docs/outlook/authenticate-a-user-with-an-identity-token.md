---
title: Authenticate a user with an identity token in an add-in
description: Learn about using the identity token provided by an Outlook add-in to implement SSO with your service.
ms.date: 03/21/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Authenticate a user with an identity token for Exchange

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

Exchange user identity tokens provide a way for your add-in to uniquely identify an add-in user. By establishing the user's identity, you can implement a single sign-on (SSO) authentication scheme for your back-end service that enables customers who are using Outlook add-ins to connect to your service without signing in. See [Exchange user identity token](authentication.md#exchange-user-identity-token) for more about when to use this token type. In this article, we'll take a look at a simplistic method of using the Exchange identity token to authenticate a user to your back-end.

> [!IMPORTANT]
> This is just a simple example of an SSO implementation. As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.

## Send the ID token with each request

The first step is for your add-in to obtain the Exchange user identity token from the server by calling [getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods). Then the add-in sends this token with every request it makes to your back-end. This could be in a header, or as part of the request body.

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

## Map the token to a user in your backend

Your back-end service can calculate a unique user ID from the token and map it to a user in your internal user system. For example, if you use a database to store users, you could add this unique ID to the user's record in your database.

### Generate a unique ID

Use a combination of the `msexchuid` and `amurl` properties. For example, you could concatenate the two values together and generate a base 64-encoded string. This value can be reliably generated from the token every time, so you can map an Exchange user identity token back to the user in your system.

### Check the user

With the unique ID generated, the next step is to check for a user in your system with that associated ID.

- If the user is found, the back-end treats the request as authenticated, and allows the request to proceed.

- If the user is not found, then the back-end returns an error indicating that the user needs to sign in. The add-in then prompts the user to sign in to the back-end using your existing authentication method. Once the user is authenticated, the Exchange user identity token is submitted with the user authentication details. The back-end can then update the user's record in your system with the unique ID.
