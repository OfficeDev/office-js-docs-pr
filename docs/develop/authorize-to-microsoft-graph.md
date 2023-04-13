---
title: Authorize to Microsoft Graph with SSO
description: Learn how users of an Office Add-in can use single sign-on (SSO) to fetch data from Microsoft Graph.
ms.date: 06/10/2022
ms.localizationpriority: medium
---

# Authorize to Microsoft Graph with SSO

Users sign in to Office using either their personal Microsoft account or their Microsoft 365 Education or work account. The best way for an Office Add-in to get authorized access to [Microsoft Graph](https://developer.microsoft.com/graph/docs) is to use the credentials from the user's Office sign on. This enables them to access their Microsoft Graph data without needing to sign in a second time.

## Add-in architecture for SSO and Microsoft Graph

In addition to hosting the pages and JavaScript of the web application, the add-in must also host, at the same [fully qualified domain name](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), one or more web APIs that will get an access token to Microsoft Graph and make requests to it.

The add-in manifest contains a **\<WebApplicationInfo\>** element that provides important Azure app registration information to Office, including the permissions to Microsoft Graph that the add-in requires.

### How it works at runtime

The following diagram shows the steps involved to sign in and access Microsoft Graph. The entire process uses OAuth 2.0 and JWT access tokens.

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="Diagram showing the SSO process." border="false":::

1. The client-side code of the add-in calls the Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)). This tells the Office host to obtain an access token for the add-in.

    If the user is not signed in, the Office host in conjunction with the Microsoft identity platform provides UI for the user to sign in and consent.

2. The Office host request an access token from the Microsoft identity platform.
3. The Microsoft identity platform returns access token *A* to the Office host. Access token *A* only provides access to the add-in's own server-side APIs. It does not provide access to Microsoft Graph.
4. The Office host returns access token *A* to the add-in's client-side code. Now the client-side code can make authenticated calls to the server-side APIs.
5. The client-side code makes an HTTP request to a web API on the server-side that requires authentication. It includes access token *A* as authorization proof. Server-side code validates access token *A*.
6. The server-side code uses the OAuth 2.0 On-Behalf-Of flow (OBO) to request a new access token with permissions to Microsoft Graph.
7. The Microsoft identity platform returns the new access token *B* with permissions to Microsoft Graph (and a refresh token, if the add-in requests *offline_access* permission). The server can optionally cache access token *B*.
8. The server-side code makes a request to a Microsoft Graph API and includes access token *B* with permissions to Microsoft Graph.
9. Microsoft Graph returns data back to the server-side code.
10. The server-side code returns the data back to the client-side code.

On subsequent requests the client code will always pass access token *A* when making authenticated calls to server-side code. The server-side code can cache token *B* so that it does not need to request it again on future API calls.

## Develop an SSO add-in that accesses Microsoft Graph

You develop an add-in that accesses Microsoft Graph just as you would any other application that uses SSO. For a thorough description, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md). The difference is that it is mandatory that the add-in have a server-side Web API.

Depending on your language and framework, libraries might be available that will simplify the server-side code you have to write. Your code should do the following:

* Validate the access token *A* every time it is passed from the client-side code. For more information, see [Validate the access token](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).
* Initiate the OAuth 2.0 On-Behalf-Of flow (OBO) with a call to the Microsoft identity platform that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret). For more information about the OBO flow, see [Microsoft identity platform and OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow).
* Optionally, after the flow completes, cache the returned access token *B* with permissions to Microsoft Graph. You'll want to do this if the add-in makes more than one call to Microsoft Graph. For more information, see [Acquire and cache tokens using the Microsoft Authentication Library (MSAL)](/azure/active-directory/develop/msal-acquire-cache-tokens)
* Create one or more Web API methods that get Microsoft Graph data by passing the (possibly cached) access token *B* to Microsoft Graph.

For examples of detailed walkthroughs and scenarios, see:

* [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md)
* [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md)
* [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md)

## Distributing SSO-enabled add-ins in Microsoft AppSource

When a Microsoft 365 admin acquires an add-in from [AppSource](https://appsource.microsoft.com), the admin can redistribute it through [Integrated Apps](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) and grant admin consent to the add-in to access Microsoft Graph scopes. It's also possible, however, for the end user to acquire the add-in directly from AppSource, in which case the user must grant consent to the add-in. This can create a potential performance problem for which we've provided a solution.

If your code passes the `allowConsentPrompt` option in the call of `getAccessToken`, like `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`, then Office can prompt the user for consent if the Microsoft identity platform reports to Office that consent has not yet been granted to the add-in. However, for security reasons, Office can only prompt the user to consent to the Microsoft Graph `profile` scope. *Office cannot prompt for consent to other Microsoft Graph scopes*, not even `User.Read`. This means that if the user grants consent on the prompt, Office returns an access token. But the attempt to exchange the access token for a new access token with additional Microsoft Graph scopes fails with error AADSTS65001, which means consent (to Microsoft Graph scopes) has not been granted.

> [!NOTE]
> The request for consent with `{ allowConsentPrompt: true }` could still fail even for the `profile` scope if the administrator has turned off end-user consent. For more information, see [Configure how end-users consent to applications using Azure Active Directory](/azure/active-directory/manage-apps/configure-user-consent).

Your code can, and should, handle this error by falling back to an alternate system of authentication, which prompts the user for consent to Microsoft Graph scopes. For code examples, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md) and [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md) and the samples they link to. The entire process requires multiple round trips to the Microsoft identity platform. To avoid this performance penalty, include the `forMSGraphAccess` option in the call of `getAccessToken`; for example, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`. This signals to Office that your add-in needs Microsoft Graph scopes. Office will ask the Microsoft identity platform to verify that consent to Microsoft Graph scopes has already been granted to the add-in. If it has, the access token is returned. If it hasn't, then the call of `getAccessToken` returns error 13012. Your code can handle this error by falling back to an alternate system of authentication immediately, without making a doomed attempt to exchange tokens with the Microsoft identity platform.

As a best practice, always pass `forMSGraphAccess` to `getAccessToken` when your add-in will be distributed in AppSource and needs Microsoft Graph scopes.

## Details on SSO with an Outlook add-in

If you develop an Outlook add-in that uses SSO and you sideload it for testing, Office will *always* return error 13012 when `forMSGraphAccess` is passed to `getAccessToken` even if administrator consent has been granted. For this reason, you should comment out the `forMSGraphAccess` option **when developing** an Outlook add-in. Be sure to uncomment the option when you deploy for production. The bogus 13012 only happens when you are sideloading in Outlook.

For Outlook add-ins, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## See also

* [OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Microsoft identity platform and OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
