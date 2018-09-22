---
title: Enable single sign-on for Office Add-ins
description: ''
ms.date: 04/10/2018
---

# Enable single sign-on for Office Add-ins (preview)

Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account. You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.


![An image showing the sign-in process for an add-in](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).
> To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.
> If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

For users, this makes running your add-in a smooth experience that involves signing in only once. For developers, this means that your add-in does not have to maintain it's own user tables with encrypted passwords.

### How it works at runtime

The following diagram shows how the SSO process works.

![A diagram that shows the SSO process](../images/sso-overview-diagram.png)

1. In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference). This tells the Office host application to obtain an access token to the add-in. See [Example access token](#example-access-token).
2. If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.
3. If this is the first time the current user has used your add-in, he or she is prompted to consent.
4. The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.
5. Azure AD sends the add-in token to the Office host application.
6. The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.
7. JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address. 
8. Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences. Alternatively, the access token itself could be sent to the server-side for parsing and validation there. 

## Develop an SSO add-in

This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:

* [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md)
* [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md)

### Create the service application

Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:

* Get a client ID and secret for the add-in.
* Specify the permissions that your add-in needs to AAD v. 2.0 endpoint (and optionally to Microsoft Graph). The "profile" permission is always needed.
* Grant the Office host application trust to the add-in.
* Preauthorize the Office host application to the add-in with the default permission *access_as_user*.

For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).

### Configure the add-in

Add new markup to the add-in manifest:

* **WebApplicationInfo** - The parent of the following elements.
* **Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in. See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).
* **Resource** - The URL of the add-in.
* **Scopes** - The parent of one or more **Scope** elements.
* **Scope** - Specifies a permission that the add-in needs to AAD. The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph. If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`. Libraries that you use in your code to access Microsoft Graph may need additional permissions. For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission. For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).

For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.

The following is an example of the markup:

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### Add client-side code

Add JavaScript to the add-in to:

* Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).

* Parse the access token or pass it to the add-in’s server-side code. 

Here's a simple example of a call to `getAccessTokenAsync`. 

> [!NOTE]
> This example handles only one kind of error explicitly. For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

Here's a simple example of passing the add-in token to the server-side. The token is included as an `Authorization` header when sending a request back to the server-side. This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### When to call the method

If your add-in cannot be used when no user is logged into Office, then you should call `getAccessTokenAsync` *when the add-in launches*.

If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*. There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v. 2.0 endpoint whenever `getAccessTokenAsync` is called. So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.

### Add server-side code

In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there. Some server-side tasks your add-in could do:

* Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base. (See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.
* Get Microsoft Graph data. Your server-side code should do the following:

    * Validate the access token (see **Validate the access token** below).
    * Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret). In this context, the access token is called the bootstrap token.
    * Cache the new access token that is returned from the on-behalf-of flow.
    * Get data from Microsoft Graph by using the new token.

 For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).

#### Validate the access token

Once the Web API receives the access token, it must validate it before using it. The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows. There are a number of libraries available that can handle JWT validation, but the basics include:

- Checking that the token is well-formed
- Checking that the token was issued by the intended authority
- Checking that the token is targeted to the Web API

Keep in mind the following guidelines when validating the token:

- Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`. The `iss` claim in the token should start with this value.
- The token's `aud` parameter will be set to the application ID of the add-in's registration.
- The token's `scp` parameter will be set to `access_as_user`.

#### Using the SSO token as an identity

If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity. The following claims in the token relate to identity.

- `name` - The user's display name.
- `preferred_username` - The user's email address.
- `oid` - A GUID representing the ID of the user in the Azure Active Directory.
- `tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.

Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.

For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database. Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.

### Example access token

The following is a typical decoded payload of an access token. For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",         
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",         
    iat: 1521143967,         
    nbf: 1521143967,         
    exp: 1521147867,         
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",         
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",         
    azpacr: "0",         
    e_exp: 262800,         
    name: "Mila Nikolova",         
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",         
    preferred_username: "milan@contoso.com",         
    scp: "access_as_user",         
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",         
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",         
    uti: "MICAQyhrH02ov54bCtIDAA",         
    ver: "2.0"
}
```

## Using SSO with an Outlook add-in

There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in. Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

## SSO API reference

### getAccessTokenAsync

The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application. Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. This enables add-ins to identify users. Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.

<table><tr><td>Hosts</td><td>Excel, OneNote, Outlook, PowerPoint, Word</td></tr>

 <tr><td>Requirement sets</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### Parameters

`options` - Optional. Accepts an `AuthOptions` object (see below) to define sign-on behaviors.

`callback` - Optional. Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph. If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v. 2.0-formatted access token.

The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v. 2.0 with the `getAccessTokenAsync` method.

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has 
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor 
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development 
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" 
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should 
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```



