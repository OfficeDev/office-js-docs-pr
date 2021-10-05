---
title: Overview of authentication and authorization in Office Add-ins
description: 'Learn how authentication and authorization works in Office Add-ins.'
ms.date: 09/20/2021
ms.localizationpriority: high
---

# Overview of authentication and authorization in Office Add-ins

Office Add-ins allow anonymous access by default, but you can require users to authenticate to use your add-in. For example, you can require that your users sign in with a Microsoft account, a Microsoft 365 Education or work account, or other common account. This task is called user authentication because it enables the add-in to know who the user is.

Your add-in can also get the user's consent to access their Microsoft Graph data (such as their Microsoft 365 profile, OneDrive files, and SharePoint data) or to data in other external sources such as Google, Facebook, LinkedIn, SalesForce, and GitHub. This task is called add-in (or app) authorization, because it is the *add-in* that is being authorized, not the user.

## Key resources for authentication and authorization

Office Add-ins rely on the Microsoft identity platform to perform authentication and authorization. This documentation explains details specific to Office Add-ins to implement authentication and authorization. This documentation does not cover general security concepts such as OAuth flows, token caching, or identity management. This documentation also does not document anything specific to Microsoft Azure or the Microsoft identity platform. We recommend you refer to the following resources if you need more information in those areas.

### Microsoft identity platform

Office Add-ins primarily depend on the Microsoft identity platform to handle authentication and authorization. We'll include links to relevant resources in the Microsoft identity platform documentation when additional details are needed. When you get stuck programming something related to the Microsoft identity platform, or need help on a concept, refer to the [Microsoft identity platform](/azure/active-directory/develop) documentation for more information.

### OAuth 2.0 and OpenID connect

You'll see OAuth 2.0 and OpenID connect mentioned in this documentation. These are standard protocols used throughout the industry that Microsoft identity platform, and Office, rely upon for authentication and authorization workflows. For more information, see [OAuth 2.0 and OpenID Connect protocols on the Microsoft identity platform](/azure/active-directory/develop/active-directory-v2-protocols)

## Authentication and authorization approach

You have a choice of two ways to accomplish authentication and authorization in your Office Add-in.

- **single sign-on (SSO) through Office**: Your add-in can access the authenticated user from Office to avoid having the user sign in twice (to Office, and then also your add-in). Optionally, your add-in can also use the user's Office sign-in to authorize your add-in to [Microsoft Graph](/graph) or other Microsoft 365 services. (Non-Microsoft sources are not accessible through this approach.)
- **Azure AD sign-in**: Your add-in can sign in users using the Microsoft identity platform as the authentication provider. Once you have signed in the user, you can then use the Microsoft identity platform to authorize the add-in to [Microsoft Graph](/graph) or other Microsoft 365 services. This approach is used when SSO through Office is unavailable. Also, there are scenarios in which you want to have your users sign in to your add-in separately even when SSO is available; for example, if you want them to have the option of signing in to the add-in with a different ID from the one with which they are currently signed in to Office.

Use the following table to find the guidance you'll need based on the required resources for your add-in, and the approach you want to use.

|Required resources  | Approach  | Guidance |
|---------|---------|---------|
|User's identity | SSO | Use the [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_) method and use the returned token as an ID token. For more information see [Get the user's information through SSO](#get-the-user's-information-through-sso). |
|User's identity and Microsoft Graph access | SSO | Use the [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_) method and use the returned token in the on-behalf-of flow to get a new access token to Microsoft Graph. For more information see [Access Microsoft Graph through SSO](#access-microsoft-Graph-through-sso). |
|User's identity | Azure AD sign-in | Authenticate as you would in any web app, but use the Dialog API to host the sign-in page. For more information, see [Authenticate with the Microsoft identity platform](#authenticate-with-the-microsoft-identity-platform).     |
|User's identity and Microsoft Graph access | Azure AD sign-in | Use Azure AD to get an access token to Graph, but use the dialog API to host the sign-in page. For more information, see [Access to Microsoft Graph without SSO](#access-to-microsoft-graph-without-sss). |
|User's identity and non-Microsoft data     | Separate sign-in        | Get authorization to the external source as you would in any web app, but you may need to use the dialog API to host the login page. See [Access to non-Microsoft data sources](#access-to-non-microsoft-data-sources). |

## SSO scenarios

Using SSO is convenient for the user because they only have to sign in once to Office. SSO is not supported on all versions of Office, so you'll still need to implement an alternative sign-in approach, such as using the Microsoft identity platform. For more information on supported versions, see [Identity API requirement sets](../reference/requirement-sets/identity-api-requirement-sets)

### Get the user's identity through SSO

Often your add-in only needs the user's identity. For example, you may just want personalize your add-in and display the user's name on the task pane. Or you might want a unique ID to associate the user with their data in your database. This can be accomplished by just getting the access token for the user from Office.

To get the user's identity through SSO, call the [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_) method. The method returns an access token containing several claims that are unique to the current signed in user, including `preferred_username`, `name`, `sub`, and `oid`. For more information on these properties, see [Microsoft identity platform access tokens](/azure/active-directory/develop/access-tokens#payload-claims). For an example of a one of these tokens, see [Example access token](sso-in-office-add-ins.md#example-access-token).

If the user is not signed in, Office will open a dialog box and use the Microsoft identity platform to request the user to sign in. Then the method will return an access token, or throw an error if unable to sign in the user.

In a scenario where you need to store data for the user, refer to [Microsoft identity platform ID tokens](azure/active-directory/develop/id-tokens) for information about how to get a value from the token to uniquely identify the user. Then use that value to look up the user in a user table or user database that you maintain. Use the database to store user-relative information such as the user's preferences or the state of the user's account. Since you are using SSO, your users don't sign-in separately to your add-in, so you do not need to store a password for the user.

Before you begin implementing user authentication with SSO, be sure that you are thoroughly familiar with the article [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md). Note also the following samples.

- [Office Add-in NodeJS SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO), especially the file [ssoAuthES6.js](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/public/javascripts/ssoAuthES6.js).
- [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).

These samples, however, do not use the token as an ID token. They use it to get access to Microsoft Graph with the **On Behalf Of** flow.

### Access Microsoft Graph through SSO

In some scenarios not only do you need the user's information, but you also need to access [Microsoft Graph](/graph) resources on behalf of the user. For example, you may need to send an email, or create a chat in Teams on behalf of the user. These actions, and more, can be accomplished through Microsoft Graph. You'll need to follow these steps:

1. Get the access token for the current user through SSO by calling [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_).
1. Use the on-behalf-of flow to exchange the access token for a new access token containing claims that allow your add-in to call Microsoft Graph.

To use SSO to access Microsoft Graph, your add-in in a task pane or function file calls the [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_) method. If the user is not signed in, Office will open a dialog box and navigate it to the Azure Active Directory login page. After the user signs in, or if the user is already signed in, the method returns an access token. The token is a bootstrap token in the **On Behalf Of** flow. Specifically, it has a `scope` claim with the value `access_as_user`. For guidance about the claims in the token, see [Microsoft identity platform access tokens](/azure/active-directory/develop/access-tokens#payload-claims). For an example of a one of these tokens, see [Example access token](sso-in-office-add-ins.md#example-access-token).

After your code obtains the token, it uses it in the **On Behalf Of** flow to obtain a second token: an access token to Microsoft Graph.

Before you begin implementing Office SSO, be sure that you are thoroughly familiar with these two articles.

- [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md)
- [Authorize to Microsoft Graph with SSO](authorize-to-microsoft-graph.md)

You should also read at least one of the walkthrough articles listed here. Even if you don't carry out the steps, these contain valuable information about how you implement Office SSO and the **On Behalf Of** flow.

- [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md)
- [Create an Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md)

Note also the following samples.

- [Office Add-in NodeJS SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)

## Non-SSO scenarios

In some scenarios you may not want to use SSO. For example, you may need to authenticate using a different identity provider than the Microsoft identity platform. Also SSO is not supported in all scenarios. For example, older versions of Office don't support SSO. You should implement an alternate authentication system that your add-in can fall back to in certain error situations.

### Authenticate with the Microsoft identity platform

You can authenticate a user in an Office Add-in with the Microsoft identity platform as you would in other web applications. You should use this approach as the alternate authentication system when SSO doesn't work.

It's important to note that the Microsoft identity platform does not allow its sign in page to open in an iframe. When an Office Add-in is running on *Office on the web*, the task pane is an iframe. This means that you'll need to open the sign in page by using a dialog box opened with the Office dialog API. This affects how you use authentication helper libraries. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

For information about implementing authentication with the Microsoft identity platform, see [Microsoft identity platform (v2.0) overview](/azure/active-directory/develop/v2-overview). The documentaiton contains many tutorials and guides, as well as links to relevant samples and libraries. As explained in [Authentication with the Office dialog API](auth-with-office-dialog-api.md), you may need to adjust the code in the samples to run in the Office dialog box.

### Access to Microsoft Graph without SSO

You can get authorization to Microsoft Graph data for your add-in by obtaining an access token to Microsoft Graph from the Microsoft identity platform. You can do this without relying on Office SSO (or if SSO failed or is not supported). For more information, see [Access to Microsoft Graph without SSO](authorize-to-microsoft-graph-without-sso.md) which has more details and links to samples.

## Access to non-Microsoft data sources

Popular online services, including Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in. For an overview of the ways that your add-in can do this, see [Authorize external services in your Office Add-in](auth-external-add-ins.md).

> [!IMPORTANT]
> Before you begin coding, find out if the data source allows its sign in page to open in an iframe. When an Office Add-in is running on *Office on the web*, the task pane is an iframe. If the data source does not allow its sign in page to open in an iframe, then you'll need to open the sign in page in a dialog box opened with the Office dialog API. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

## See also

- [Microsoft identity platform documentation](/azure/active-directory/develop/)
- [Microsoft identity platform access tokens](https://docs.microsoft.com/en-us/azure/active-directory/develop/access-tokens)
- [OAuth 2.0 and OpenID Connect protocols on the Microsoft identity platform](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols)
- [Microsoft identity platform and OAuth 2.0 On-Behalf-Of flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
- [JSON web token (JWT)](https://en.wikipedia.org/wiki/JSON_Web_Token)
- [JSON web token viewer](https://jwt.ms/)
