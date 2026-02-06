---
title: Overview of authentication and authorization in Office Add-ins
description: Learn how authentication and authorization works in Office Add-ins.
ms.date: 12/25/2025
ms.localizationpriority: high
---

# Overview of authentication and authorization in Office Add-ins

Office Add-ins allow anonymous access by default, but you can require users to sign in to use your add-in with a Microsoft account, a Microsoft 365 Education or work account, or other common account. This task is called user authentication because it enables the add-in to know who the user is.

Your add-in can also get the user's consent to access their Microsoft Graph data (such as their Microsoft 365 profile, OneDrive files, and SharePoint data) or data in other external sources such as Google, Facebook, LinkedIn, SalesForce, and GitHub. This task is called add-in (or app) authorization, because it's the *add-in* that is being authorized, not the user.

## Key resources for authentication and authorization

This documentation describes how to build and configure Office Add-ins to support authentication and authorization. Some concepts and security technologies referenced are outside the scope of this content. For example, general security topics such as OAuth flows, token caching, and identity management are not covered. This documentation also does not include guidance specific to Microsoft Entra ID or the Microsoft identity platform. Use the following resources for additional information on those topics.

- [Microsoft identity platform documentation](/entra/identity-platform/)
- [Microsoft identity platform support and help options for developers](/entra/identity-platform/developer-support-help-options)
- [OAuth 2.0 and OIDC authentication flow in the Microsoft identity platform](/entra/identity-platform/v2-protocols)

## Single Sign-on (SSO)

Single sign-on (SSO) improves the user experience by allowing users to sign in once to Office. Users aren’t required to sign in again when interacting with the add-in.

To get started with SSO, see [Enable single sign-on in an Office Add-in with nested app authentication](enable-nested-app-authentication-in-your-add-in.md)

### Legacy SSO

Office.js also has a `getAccessToken` API you can use to get an access token for the signed in user. This approach is considered legacy but is still supported. If you want to use Office.js for SSO, or are maintaining an add-in that was built using Office.js for SSO, see [Enable single sign-on (SSO) in an Office Add-in](sso-in-office-add-ins.md).

## Non-SSO scenarios

In some scenarios, you may not want to use SSO. For example, you may need to authenticate using a different identity provider than the Microsoft identity platform. Also, SSO isn't supported in all scenarios. For example, older versions of Office don't support SSO. In this case, you'd need to fall back to an alternate authentication system for your add-in.

### Authenticate with the Microsoft identity platform

Your add-in can sign in users using the [Microsoft identity platform](/azure/active-directory/develop) as the authentication provider. Once you've signed in the user, you can then use the Microsoft identity platform to authorize the add-in to [Microsoft Graph](/graph) or other services managed by Microsoft. Use this approach as an alternate sign-in method when SSO through Office is unavailable. Also, there are scenarios in which you want to have your users sign in to your add-in separately even when SSO is available; for example, if you want them to have the option of signing in to the add-in with a different ID from the one with which they're currently signed in to Office.

It's important to note that the Microsoft identity platform doesn't allow its sign-in page to open in an iframe. When an Office Add-in is running in *Office on the web*, the task pane is an iframe. This means that you'll need to open the sign-in page by using a dialog box opened with the Office dialog API. This affects how you use authentication helper libraries. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

For information about implementing authentication with the Microsoft identity platform, see [What is the Microsoft identity platform?](/entra/identity-platform/v2-overview) The documentation contains many tutorials and guides, as well as links to relevant samples and libraries. As explained in [Authentication with the Office dialog API](auth-with-office-dialog-api.md), you may need to adjust the code in the samples to run in the Office dialog box.

### Access to Microsoft Graph without SSO

You can get authorization to Microsoft Graph data for your add-in by obtaining an access token to Microsoft Graph from the Microsoft identity platform. You can do this without relying on SSO through Office (or if SSO failed or isn't supported). For more information, see [Access to Microsoft Graph without SSO](authorize-to-microsoft-graph-without-sso.md) which has more details and links to samples.

### Access to non-Microsoft data sources

Popular online services, including Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in. For an overview of the ways that your add-in can do this, see [Authorization with non-Microsoft identity providers](auth-external-add-ins.md).

> [!IMPORTANT]
> Before you begin coding, verify whether the data source supports opening its sign-in page in an iframe. In *Office on the web*, the task pane runs inside an iframe. If the data source doesn't allow its sign-in page to load in an iframe, open the sign-in page in a dialog box using the Office dialog API. For more information, see [Authentication with the Office dialog API](auth-with-office-dialog-api.md).

## See also

- [Microsoft identity platform documentation](/entra/identity-platform/)
- [ID tokens in the Microsoft identity platform](/entra/identity-platform/id-tokens)
- [Access tokens in the Microsoft identity platform](/entra/identity-platform/access-tokens)
- [OAuth 2.0 and OIDC authentication flow in the Microsoft identity platform](/entra/identity-platform/v2-protocols)
- [JSON web token (JWT)](https://en.wikipedia.org/wiki/JSON_Web_Token)
- [JSON web token viewer](https://jwt.ms/)
