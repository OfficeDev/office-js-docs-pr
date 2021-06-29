---
title: Authentication options in Outlook add-ins
description: Outlook add-ins provide a number of different methods to authenticate, depending on your specific scenario.
ms.date: 06/22/2021
localization_priority: Priority
---

# Authentication options in Outlook add-ins

Your Outlook add-in can access information from anywhere on the Internet, whether from the server that hosts the add-in, from your internal network, or from somewhere else in the cloud. If that information is protected, your add-in needs a way to authenticate your user. Outlook add-ins provide a number of different methods to authenticate, depending on your specific scenario.

## Single sign-on access token

Single sign-on access tokens provide a seamless way for your add-in to authenticate and obtain access tokens to call the [Microsoft Graph API](/graph/overview). This capability reduces friction since the user is not required to enter their credentials.

> [!NOTE]
> The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).
> If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Consider using SSO access tokens if your add-in:

- Is used primarily by Microsoft 365 users
- Needs access to:
  - Microsoft services that are exposed as part of Microsoft Graph
  - A non-Microsoft service that you control

The SSO authentication method uses the [OAuth2 On-Behalf-Of flow provided by Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of). It requires that the add-in register in the [Application Registration Portal](https://apps.dev.microsoft.com/) and specify any required Microsoft Graph scopes in its manifest.

Using this method, your add-in can obtain an access token scoped to your server back-end API. The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API. At that point your server can:

- Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API
- Use the identity information in the token to establish the user's identity and authenticate to your own back-end services

For a more detailed overview, see the [full overview of the SSO authentication method](../develop/sso-in-office-add-ins.md).

For details on using the SSO token in an Outlook add-in, see [Authenticate a user with an single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md).

For a sample add-in that uses the SSO token, see [Outlook Add-in SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO).

## Exchange user identity token

Exchange user identity tokens provide a way for your add-in to establish the identity of the user. By verifying the user's identity, you can then perform a one-time authentication into your back-end system, then accept the user identity token as an authorization for future requests. Use the Exchange user identity token:

- When the add-in is used primarily by Exchange on-premises users.
- When the add-in needs access to a non-Microsoft service that you control.
- As a fallback authentication when the add-in is running on a version of Office that doesn't support SSO.

Your add-in can call [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#getuseridentitytokenasync-callback--usercontext-) to get Exchange user identity tokens. For details on using these tokens, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).

## Access tokens obtained via OAuth2 flows

Add-ins can also access third-party services that support OAuth2 for authorization. Consider using OAuth2 tokens if your add-in:

- Needs access to a third-party service outside of your control

Using this method, your add-in prompts the user to sign-in to the service either by using the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method to initialize the OAuth2 flow, or by using the [office-js-helpers library](https://github.com/OfficeDev/office-js-helpers) to the OAuth2 Implicit flow.

## Callback tokens

Callback tokens provide access to the user's mailbox from your server back-end, either using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange), or the [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api). Consider using callback tokens if your add-in:

- Needs access to the user's mailbox from your server back-end.

Add-ins obtain callback tokens using one of the [getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) methods. The level of access is controlled by the permissions specified in the add-in manifest.
