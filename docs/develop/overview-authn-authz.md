---
title: Overview of authentication and authorization in Office Add-ins
description: Learn how authentication and authorization works in Office Add-ins.
ms.date: 06/23/2023
ms.localizationpriority: high
---

# Overview of authentication and authorization in Office Add-ins

Office Add-ins allow anonymous access by default, but you can require users to sign in to use your add-in with a Microsoft account, a Microsoft 365 Education or work account, or other common account. This task is called user authentication because it enables the add-in to know who the user is.

Your add-in can also get the user's consent to access their Microsoft Graph data (such as their Microsoft 365 profile, OneDrive files, and SharePoint data) or data in other external sources such as Google, Facebook, LinkedIn, SalesForce, and GitHub. This task is called add-in (or app) authorization, because it's the *add-in* that is being authorized, not the user.

## Key resources for authentication and authorization

This documentation explains how to build and configure Office Add-ins to successfully implement authentication and authorization. However, many concepts and security technologies mentioned are outside the scope of this documentation. For example, general security concepts such as OAuth flows, token caching, or identity management are not explained here. This documentation also doesn't document anything specific to Microsoft Azure or the Microsoft identity platform. We recommend you refer to the following resources if you need more information in those areas.

- [Microsoft identity platform](/azure/active-directory/develop)
- [Microsoft identity platform support and help options for developers](/azure/active-directory/develop/developer-support-help-options)
- [OAuth 2.0 and OpenID Connect protocols on the Microsoft identity platform](/azure/active-directory/develop/active-directory-v2-protocols)

## SSO scenarios

Using Single Sign-on (SSO) is convenient for the user because they only have to sign in once to Office. They don't need to sign in separately to your add-in. SSO isn't supported on all versions of Office, so you'll still need to implement an alternative sign-in approach, by [using the Microsoft identity platform](#authenticate-with-the-microsoft-identity-platform). For more information on supported Office versions, see [Identity API requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets)

### Get the user's identity through SSO

Often your add-in only needs the user's identity. For example, you may just want to personalize your add-in and display the user's name on the task pane. Or you might want a unique ID to associate the user with their data in your database. This can be accomplished by just getting the access token for the user from Office.

To get the user's identity through SSO, call the [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) method. The method returns an access token that is also an identity token containing several claims that are unique to the current signed in user, including `preferred_username`, `name`, `sub`, and `oid`. For more information on these properties, see [Microsoft identity platform ID tokens](/azure/active-directory/develop/id-tokens). For an example of the token returned by [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)), see [Example access token](sso-in-office-add-ins.md#example-access-token).

If the user isn't signed in, Office will open a dialog box and use the Microsoft identity platform to request the user to sign in. Then the method will return an access token, or throw an error if unable to sign in the user.

In a scenario where you need to store data for the user, refer to [Microsoft identity platform ID tokens](/azure/active-directory/develop/id-tokens) for information about how to get a value from the token to uniquely identify the user. Use that value to look up the user in a user table or user database that you maintain. Use the database to store user-relative information such as the user's preferences or the state of the user's account. Since you're using SSO, your users don't sign-in separately to your add-in, so you don't need to store a password for the user.

Before you begin implementing user authentication with SSO, be sure that you're thoroughly familiar with the article [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md).

### Access your Web APIs through SSO

If your add-in has server-side APIs that require an authorized user, call the [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) method to get an access token. The access token provides access to your own web server (configured through an [app registration in Microsoft Entra ID](register-sso-add-in-aad-v2.md)). When you call APIs on your web server, you also pass the access token to authorize the user.

The following code shows how to construct an HTTPS GET request to the add-in's web server API to get some data. The code runs on the client side, such as in a task pane. It first gets the access token by calling `getAccessToken`. Then it constructs an AJAX call with the correct authorization header and URL for the server API.

```javascript
function getOneDriveFileNames() {

    let accessToken = await Office.auth.getAccessToken();

    $.ajax({
        url: "/api/data",
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET"
    })
        .done(function (result) {
            //... work with data from the result...
        });
}
```

The following code shows an example /api/data handler for the REST call from the previous code example. The code is ASP.NET code running on a web server. The `[Authorize]` attribute will require that a valid access token is passed from the client, or it'll return an error to the client.

```csharp
    [Authorize]
    // GET api/data
    public async Task<HttpResponseMessage> Get()
    {
        //... obtain and return data to the client-side code...
    }
```

### Access Microsoft Graph through SSO

In some scenarios, not only do you need the user's identity, but you also need to access [Microsoft Graph](/graph) resources on behalf of the user. For example, you may need to send an email, or create a chat in Teams on behalf of the user. These actions and more can be accomplished through Microsoft Graph. You'll need to follow these steps:

1. Get the access token for the current user through SSO by calling [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)). If the user isn't signed in, Office will open a dialog box and sign in the user with the Microsoft identity platform. After the user signs in, or if the user is already signed in, the method returns an access token.
1. Pass the access token to your server-side code.
1. On the server-side, use the [OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) to exchange the access token for a new access token containing the necessary delegated user identity and permissions to call Microsoft Graph.

> [!NOTE]
> For best security to avoid leaking the access token, always perform the On-Behalf-Of flow on the server-side. Call Microsoft Graph APIs from your server, not the client. Don't return the access token to the client-side code.

Before you begin implementing SSO to access Microsoft Graph in your add-in, be sure that you're thoroughly familiar with the following articles.

- [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md)
- [Authorize to Microsoft Graph with SSO](authorize-to-microsoft-graph.md)

You should also read at least one of the following articles that'll walk you through building an Office Add-in to use SSO and access Microsoft Graph. Even if you don't carry out the steps, they contain valuable information about how you implement SSO and the On-Behalf-Of flow.

- [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md) which walks you through the sample at [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).
- [Create an Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md) which walks you through the sample at [Office Add-in NodeJS SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

## Non-SSO scenarios

In some scenarios, you may not want to use SSO. For example, you may need to authenticate using a different identity provider than the Microsoft identity platform. Also, SSO isn't supported in all scenarios. For example, older versions of Office don't support SSO. In this case, you'd need to fall back to an alternate authentication system for your add-in.

### Authenticate with the Microsoft identity platform

Your add-in can sign in users using the [Microsoft identity platform](/azure/active-directory/develop) as the authentication provider. Once you've signed in the user, you can then use the Microsoft identity platform to authorize the add-in to [Microsoft Graph](/graph) or other services managed by Microsoft. Use this approach as an alternate sign-in method when SSO through Office is unavailable. Also, there are scenarios in which you want to have your users sign in to your add-in separately even when SSO is available; for example, if you want them to have the option of signing in to the add-in with a different ID from the one with which they're currently signed in to Office.

It's important to note that the Microsoft identity platform doesn't allow its sign-in page to open in an iframe. When an Office Add-in is running in *Office on the web*, the task pane is an iframe. This means that you'll need to open the sign-in page by using a dialog box opened with the Office dialog API. This affects how you use authentication helper libraries. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

For information about implementing authentication with the Microsoft identity platform, see [Microsoft identity platform (v2.0) overview](/azure/active-directory/develop/v2-overview). The documentation contains many tutorials and guides, as well as links to relevant samples and libraries. As explained in [Authentication with the Office dialog API](auth-with-office-dialog-api.md), you may need to adjust the code in the samples to run in the Office dialog box.

### Access to Microsoft Graph without SSO

You can get authorization to Microsoft Graph data for your add-in by obtaining an access token to Microsoft Graph from the Microsoft identity platform. You can do this without relying on SSO through Office (or if SSO failed or isn't supported). For more information, see [Access to Microsoft Graph without SSO](authorize-to-microsoft-graph-without-sso.md) which has more details and links to samples.

### Access to non-Microsoft data sources

Popular online services, including Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in. For an overview of the ways that your add-in can do this, see [Authorize external services in your Office Add-in](auth-external-add-ins.md).

> [!IMPORTANT]
> Before you begin coding, find out if the data source allows its sign-in page to open in an iframe. When an Office Add-in is running in *Office on the web*, the task pane is an iframe. If the data source doesn't allow its sign-in page to open in an iframe, then you'll need to open the sign-in page in a dialog box opened with the Office dialog API. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

## See also

- [Microsoft identity platform documentation](/azure/active-directory/develop/)
- [Microsoft identity platform access tokens](/azure/active-directory/develop/access-tokens)
- [OAuth 2.0 and OpenID Connect protocols on the Microsoft identity platform](/azure/active-directory/develop/active-directory-v2-protocols)
- [Microsoft identity platform and OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
- [JSON web token (JWT)](https://en.wikipedia.org/wiki/JSON_Web_Token)
- [JSON web token viewer](https://jwt.ms/)
