---
title: Authentication options in Outlook add-ins
description: Outlook add-ins provide a number of different methods to authenticate, depending on your specific scenario.
ms.date: 10/29/2024
ms.topic: overview
ms.localizationpriority: high
---

# Authentication options in Outlook add-ins

Your Outlook add-in can access information from anywhere on the Internet, whether from the server that hosts the add-in, from your internal network, or from somewhere else in the cloud. If that information is protected, your add-in needs a way to authenticate your user. Outlook add-ins provide a number of different methods to authenticate, depending on your specific scenario.

## Single sign-on access token using OBO flow

Single sign-on (SSO) access tokens provide a seamless way for your Outlook add-in to authenticate and obtain access tokens to call the [Microsoft Graph API](/graph/overview). This capability reduces friction since the user is not required to enter their credentials. You can use the on-behalf-of flow with a middle-tier server, or nested app authentication (described in the next section).

> [!NOTE]
> SSO using the OBO flow is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about support, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
> If you're working with an Outlook add-in, be sure to enable modern authentication for the Microsoft 365 tenancy. For information about how to do this, see [Enable or disable modern authentication for Outlook in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online).

Consider using SSO access tokens if your add-in:

- Is used primarily by Microsoft 365 users
- Needs access to:
  - Microsoft services that are exposed as part of Microsoft Graph
  - A non-Microsoft service that you control

The SSO authentication method uses the [OAuth2 On-Behalf-Of flow provided by Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of). It requires that the add-in register in the [Application Registration Portal](https://apps.dev.microsoft.com/) and specify any required Microsoft Graph scopes in the add-in manifest, if the add-in only manifest is being used. If the add-in is using the [unified manifest for Microsoft 365](../develop/json-manifest-overview.md), there is some manifest configuration, but Microsoft Graph scopes aren't specified. Instead, the needed scopes are specified at runtime in a call to fetch a token to Microsoft Graph.

Using this method, your add-in can obtain an access token scoped to your server back-end API. The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API. At that point your server can:

- Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API
- Use the identity information in the token to establish the user's identity and authenticate to your own back-end services

For a more detailed overview, see the [full overview of the SSO authentication method](../develop/sso-in-office-add-ins.md).

For details on using the SSO token in an Outlook add-in, see [Authenticate a user with an single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md).

For a sample add-in that uses the SSO token, see [Outlook Add-in SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).

## Single sign-on access token using nested app authentication

Nested App Authentication (NAA) enables Single Sign-On (SSO) for Office Add-ins running in the context of native Office applications. Compared with the on-behalf-of flow used with Office.js and `getAccessToken()`, NAA provides greater flexibility in app architecture, enabling the creation of rich, client-driven applications. NAA makes handling SSO simpler for your add-in code. NAA enables you to make Microsoft Graph calls from your add-in client code as an SPA without the need for a middle-tier server. There’s no need to use Office.js APIs as NAA is provided by the MSAL.js library.

To enable your Outlook add-in to use NAA, see [Enable SSO in an Office Add-in using nested app authentication (preview)](../develop/enable-nested-app-authentication-in-your-add-in.md). NAA works the same across all Office Add-ins.

## Exchange user identity token

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

Exchange user identity tokens provide a way for your add-in to establish the identity of the user. By verifying the user's identity, you can then perform a one-time authentication into your back-end system, then accept the user identity token as an authorization for future requests. Use the Exchange user identity token:

- When the add-in is used primarily by Exchange on-premises users.
- When the add-in needs access to a non-Microsoft service that you control.
- As a fallback authentication when the add-in is running on a version of Office that doesn't support SSO.

Your add-in can call [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getuseridentitytokenasync-member(1)) to get Exchange user identity tokens. For details on using these tokens, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).

## Access tokens obtained via OAuth2 flows

Add-ins can also access services from Microsoft and others that support OAuth2 for authorization. Consider using OAuth2 tokens if your add-in:

- Needs access to a service outside of your control.

Using this method, your add-in prompts the user to sign-in to the service either by using the [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) method to initialize the OAuth2 flow.

## Callback tokens

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

Callback tokens provide access to the user's mailbox from your server back-end, either using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange), or the [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api). Consider using callback tokens if your add-in:

- Needs access to the user's mailbox from your server back-end.

Add-ins obtain callback tokens using one of the [getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) methods. The level of access is controlled by the permissions specified in the add-in manifest.

## See also

- [Nested app auth requirement set](/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets)
- [Enable SSO in an Office Add-in using nested app authentication (preview)](../develop/enable-nested-app-authentication-in-your-add-in.md)
- [Nested app authentication and Outlook legacy tokens deprecation FAQ](faq-nested-app-auth-outlook-legacy-tokens.md)
