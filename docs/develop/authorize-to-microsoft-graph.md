---
title: Authorize to Microsoft Graph with legacy Office SSO
description: Learn how users of an Office Add-in can use legacy Office single sign-on (SSO) to fetch data from Microsoft Graph.
ms.date: 05/25/2025
ms.localizationpriority: medium
---

# Authorize to Microsoft Graph with legacy Office SSO

> [!NOTE]
> This article describes legacy Office single sign-on (SSO). For a modern authentication experience with support across a wider range of platforms, use the Microsoft Authentication Library (MSAL) with nested app authentication (NAA). For more information, see [Enable single sign-on in an Office Add-in with nested app authentication](enable-nested-app-authentication-in-your-add-in.md).

Users sign in to Office using either their personal Microsoft account or their Microsoft 365 Education or work account. The best way for an Office Add-in to get authorized access to [Microsoft Graph](https://developer.microsoft.com/graph/docs) is to use the credentials from the user's Office sign on. This enables them to access their Microsoft Graph data without needing to sign in a second time.

## Add-in architecture for legacy Office SSO and Microsoft Graph

In addition to hosting the pages and JavaScript of the web application, the add-in must also host, at the same [fully qualified domain name](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), one or more web APIs that will get an access token to Microsoft Graph and make requests to it.

The add-in manifest contains a `<WebApplicationInfo>` element that provides important Azure app registration information to Office, including the permissions to Microsoft Graph that the add-in requires.

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

## Develop a legacy Office SSO add-in that accesses Microsoft Graph

You develop an add-in that accesses Microsoft Graph just as you would any other application that uses SSO. For a thorough description, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md). The difference is that it is mandatory that the add-in have a server-side Web API.

Depending on your language and framework, libraries might be available that will simplify the server-side code you have to write. Your code should do the following:

* Validate the access token *A* every time it is passed from the client-side code. For more information, see [Validate the access token](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).
* Initiate the OAuth 2.0 On-Behalf-Of flow (OBO) with a call to the Microsoft identity platform that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret). For more information about the OBO flow, see [Microsoft identity platform and OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow).
* Optionally, after the flow completes, cache the returned access token *B* with permissions to Microsoft Graph. You'll want to do this if the add-in makes more than one call to Microsoft Graph. For more information, see [Acquire and cache tokens using the Microsoft Authentication Library (MSAL)](/azure/active-directory/develop/msal-acquire-cache-tokens)
* Create one or more Web API methods that get Microsoft Graph data by passing the (possibly cached) access token *B* to Microsoft Graph.

For examples of detailed walkthroughs and scenarios, see:

* [Create a Node.js Office Add-in that uses single sign-on](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
* [Create an ASP.NET Office Add-in that uses single sign-on](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
* [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md)

## Distributing legacy Office SSO-enabled add-ins in Microsoft Marketplace

When a Microsoft 365 admin acquires an add-in from [Microsoft Marketplace](https://marketplace.microsoft.com), the admin can redistribute it through the [integrated apps portal](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) and grant admin consent to the add-in to access Microsoft Graph scopes. It's also possible, however, for the end user to acquire the add-in directly from Microsoft Marketplace, in which case the user must grant consent to the add-in. This can create a potential performance problem for which we've provided a solution.

If your code passes the `allowConsentPrompt` option in the call of `getAccessToken`, like `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`, then Office can prompt the user for consent if the Microsoft identity platform reports to Office that consent has not yet been granted to the add-in. However, for security reasons, Office can only prompt the user to consent to the Microsoft Graph `profile` scope. *Office cannot prompt for consent to other Microsoft Graph scopes*, not even `User.Read`. This means that if the user grants consent on the prompt, Office returns an access token. But the attempt to exchange the access token for a new access token with additional Microsoft Graph scopes fails with error AADSTS65001, which means consent (to Microsoft Graph scopes) has not been granted.

> [!NOTE]
> The request for consent with `{ allowConsentPrompt: true }` could still fail even for the `profile` scope if the administrator has turned off end-user consent. For more information, see [Configure how end-users consent to applications](/entra/identity/enterprise-apps/configure-user-consent).

Your code can handle this error by falling back to an alternate system of authentication that prompts the user to grant consent for Microsoft Graph scopes. For examples, see [Create a Node.js Office Add-in that uses single sign-on](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) and [Create an ASP.NET Office Add-in that uses single sign-on](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) and the samples they link to.
This flow can involve multiple round trips to the Microsoft identity platform. To reduce the performance impact, include the `forMSGraphAccess` option when calling `getAccessToken`; for example, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`. This option indicates that your add-in requires Microsoft Graph scopes. Office will ask the Microsoft identity platform to determine whether consent has already been granted. If consent exists, the access token is returned. If not, `getAccessToken` returns error **13012**. Your code can handle this error by falling back to an alternate system of authentication immediately, without attempting to exchange tokens with the Microsoft identity platform.

As a best practice, always pass `forMSGraphAccess` to `getAccessToken` when your add-in will be distributed in Microsoft Marketplace and needs Microsoft Graph scopes.

## Details on Office legacy SSO with an Outlook add-in

If you develop an Outlook add-in that uses SSO and you sideload it for testing, Office will *always* return error 13012 when `forMSGraphAccess` is passed to `getAccessToken` even if administrator consent has been granted. For this reason, you should comment out the `forMSGraphAccess` option **when developing** an Outlook add-in. Be sure to uncomment the option when you deploy for production. The bogus 13012 only happens when you are sideloading in Outlook.

For Outlook add-ins, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Enable or disable modern authentication for Outlook in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online).

[!INCLUDE [chrome-tracking-prevention](../includes/chrome-tracking-prevention.md)]

## See also

* [OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Microsoft identity platform and OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
