---
title: Authorize to Microsoft Graph with SSO
description: 'Learn how users of an Office Add-in can use single sign-on (SSO) to fetch data from Microsoft Graph.'
ms.date: 07/30/2020
localization_priority: Normal
---

# Authorize to Microsoft Graph with SSO

Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their Microsoft 365 Education or work account. The best way for an Office Add-in to get authorized access to [Microsoft Graph](https://developer.microsoft.com/graph/docs) is to use the credentials from the user's Office sign on. This enables them to access their Microsoft Graph data without needing to sign in a second time.

> [!NOTE]
> The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).
> If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).


## Add-in architecture for SSO and Microsoft Graph

In addition to hosting the pages and JavaScript of the web application, the add-in must also host, at the same [fully qualified domain name](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), one or more web APIs that will get an access token to Microsoft Graph and make requests to it.

The add-in manifest contains markup that specifies how the add-in is registered in the Azure Active Directory (Azure AD) v2.0 endpoint, and it specifies any permissions to Microsoft Graph that the add-in needs.

### How it works at runtime

The following diagram shows how the process of signing in and getting access to Microsoft Graph works.

![A diagram that shows the SSO process](../images/sso-access-to-microsoft-graph.png)

1. In the add-in, JavaScript calls a new Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-). This tells the Office client application to obtain an access token to the add-in. (Hereafter, this is called the **bootstrap access token** because it is replaced with a second token later in the process. For an example of a decoded bootstrap access token, see [Example access token](sso-in-office-add-ins.md#example-access-token).)
2. If the user is not signed in, the Office client application opens a pop-up window for the user to sign in.
3. If this is the first time the current user has used your add-in, he or she is prompted to consent.
4. The Office client application requests the **bootstrap access token** from the Azure AD v2.0 endpoint for the current user.
5. Azure AD sends the bootstrap token to the Office client application.
6. The Office client application sends the **bootstrap access token** to the add-in as part of the result object returned by the `getAccessToken` call.
7. JavaScript in the add-in makes an HTTP request to a web API that is hosted at the same fully-qualified domain as the add-in, and it includes the **bootstrap access token** as authorization proof.
8. Server-side code validates the incoming **bootstrap access token**.
9. Server-side code uses the "on behalf of" flow (defined at [OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) and the [daemon or server application to web API Azure scenario](/azure/active-directory/develop/active-directory-authentication-scenarios)) to obtain an access token for Microsoft Graph in exchange for the bootstrap access token.
10. Azure AD returns the access token to Microsoft Graph (and a refresh token, if the add-in requests *offline_access* permission) to the add-in.
11. Server-side code caches the access token to Microsoft Graph.
12. Server-side code makes requests to Microsoft Graph and includes the access token to Microsoft Graph.
13. Microsoft Graph returns data to the add-in, which can pass it on to the add-in's UI.
14. When the access token to Microsoft Graph expires, the server-side code can use its refresh token to get a new access token to Microsoft Graph.

## Develop an SSO add-in that accesses Microsoft Graph

You develop an add-in that accesses Microsoft Graph just as you would any other add-in that uses SSO. For a thorough description, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md). The difference is that it is mandatory that the add-in have a server-side Web API, and what's called the access token in that article is called the "bootstrap access token."

Depending on your language and framework, libraries might be available that will simplify the server-side code you have to write. Your code should do the following:

* Initiate the "on behalf of" flow with a call to the Azure AD v2.0 endpoint that includes the bootstrap access token, some metadata about the user, and the credentials of the add-in (its ID and secret).
* Create one or more Web API methods that get Microsoft Graph data by passing the (possibly cached) access token to Microsoft Graph.
* Optionally, before initiating the flow, validate the add-in bootstrap access token that is received from the token handler you created earlier. For more information, see [Validate the access token](sso-in-office-add-ins.md#validate-the-access-token). 
* Optionally, after the flow completes, cache the returned access token to Microsoft Graph. You'll want to do this if the add-in makes more than one call to Microsoft Graph. For more information about this flow see [Azure Active Directory v2.0 and OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> For examples of decoded access tokens for Microsoft Graph that have been obtained by the "on-behalf-of" flow, see [Azure Active Directory v2.0 and OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

For examples of detailed walkthroughs and scenarios, see:

* [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md)
* [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md)
* [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md)

## Distributing SSO-enabled add-ins in Microsoft AppSource

When a Microsoft 365 admin acquires an add-in from [AppSource](https://appsource.microsoft.com), the admin can redistribute it by [centralized deployment](../publish/centralized-deployment.md) and grant admin consent to the add-in to access Microsoft Graph scopes. It's also possible, however, for the end user to acquire the add-in directly from AppSource, in which case the user must grant consent to the add-in. This can create a potential performance problem for which we've provided a solution.

If your code passes the `allowConsentPrompt` option in the call of `getAccessToken`, like `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`, then Office can prompt the user for consent if Azure AD reports to Office that consent has not yet been granted to the add-in. However, for security reasons, Office can only prompt the user to consent to the Azure AD `profile` scope. *Office cannot prompt for consent to any Microsoft Graph scopes*, not even `User.Read`. This means that if the user grants consent on the prompt, Office will return a bootstrap token. But the attempt to exchange the bootstrap token for an access token to Microsoft Graph will fail with error AADSTS65001, which means consent (to Microsoft Graph scopes) has not been granted.

Your code can, and should, handle this error by falling back to an alternate system of authentication, which will prompt the user for consent to Microsoft Graph scopes. (For code examples, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md) and [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md) and the samples they link to.) But the entire process requires multiple round trips to Azure AD. You can avoid this performance penalty by including the `forMSGraphAccess` option in the call of `getAccessToken`; for example, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`.  This signals to Office that your add-in needs Microsoft Graph scopes. Office will ask Azure AD to verify that consent to Microsoft Graph scopes has already been granted to the add-in. If it has, the bootstrap token will be returned. If it hasn't, then the call of `getAccessToken` will return error 13012. Your code can handle this error by falling back to an alternate system of authentication immediately, without making a doomed attempt to exchange tokens with Azure AD.

As a best practice, always pass `forMSGraphAccess` to `getAccessToken` when your add-in will be distributed in AppSource and needs Microsoft Graph scopes.

> [!TIP]
> If you develop an Outlook add-in that uses SSO and you sideload it for testing, Office will *always* return error 13012 when `forMSGraphAccess` is passed to `getAccessToken` even if administrator consent has been granted. For this reason, you should comment out the `forMSGraphAccess` option **when developing** an Outlook add-in. Be sure to uncomment the option when you deploy for production. The bogus 13012 only happens when you are sideloading in Outlook.
