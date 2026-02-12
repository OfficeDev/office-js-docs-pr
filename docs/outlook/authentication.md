---
title: Authentication options in Outlook add-ins
description: Outlook add-ins provide a number of different methods to authenticate, depending on your specific scenario.
ms.date: 02/02/2026
ms.topic: overview
ms.localizationpriority: high
---

# Authentication options in Outlook add-ins

Your Outlook add-in can access information from anywhere on the Internet, whether from the server that hosts the add-in, from your internal network, or from somewhere else in the cloud. If that information is protected, your add-in needs a way to authenticate your user. Outlook add-ins provide a number of different methods to authenticate, depending on your specific scenario.

## Single sign-on access token using nested app authentication

Single sign-on (SSO) improves the user experience by allowing users to sign in once to Office. Users aren’t required to sign in again when interacting with the add-in. Nested App Authentication (NAA) enables SSO for Office Add-ins running in the context of native Office applications. NAA makes handling SSO simpler for your add-in code. NAA enables you to make Microsoft Graph calls from your add-in client code as an SPA without the need for a middle-tier server. There’s no need to use Office.js APIs as NAA is provided by the MSAL.js library.

To enable your Outlook add-in to use NAA, see [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md). For more information about support, see [Nested app auth requirement set](/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets).

### NAA samples for Outlook

- [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA)
- [Send identity claims to resources using nested app authentication (NAA) and SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA-Identity)
- [Implement SSO in events in an Outlook add-in using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Event-SSO-NAA)

## Legacy Office single sign-on using the on-behalf-of (OBO) flow

You can also use the legacy Office.js `getAccessToken` API to enable SSO in your Outlook add-in. However, this is a legacy pattern that should only be used if you're maintaining an existing add-in or need to support older Outlook clients on platforms without NAA. SSO using the OBO flow is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about support, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).

For a more detailed overview, see the [full overview of the SSO authentication method](../develop/sso-in-office-add-ins.md).

For a sample add-in that uses the SSO token, see [Outlook Add-in SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).

## Access tokens obtained via OAuth2 flows

Add-ins can also access services from Microsoft and others that support OAuth2 for authorization. Consider using OAuth2 tokens if your add-in needs access to a service outside of your control.

Using this method, your add-in prompts the user to sign in to the service using the [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) method to initialize the OAuth2 flow. For more information, see [Authenticate and authorize with the Office dialog API](../develop/auth-with-office-dialog-api.md).

## Exchange on-premises flows

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

> [!NOTE]
> Most functionality in an Exchange user identity or callback token can also be achieved by using the [Microsoft Graph mail API](/graph/outlook-mail-concept-overview).

### Exchange user identity token

Exchange user identity tokens provide a way for your add-in to establish the identity of the user. By verifying the user's identity, you can then perform a one-time authentication into your back-end system, then accept the user identity token as an authorization for future requests. Use the Exchange user identity token:

- Only for when the add-in is used by Exchange on-premises users.
- When the add-in needs access to a non-Microsoft service that you control.
- As a fallback authentication when the add-in is running on a version of Office that doesn't support SSO.

Your add-in can call [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getuseridentitytokenasync-member(1)) to get Exchange user identity tokens. For details on using these tokens, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).

### Callback tokens

Callback tokens provide access to the user's mailbox from your server back-end using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange). Consider using callback tokens if your add-in needs access to the user's mailbox from your server back-end.

Add-ins obtain callback tokens using one of the [getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) methods. The level of access is controlled by the permissions specified in the add-in manifest.

## See also

- [Nested app auth requirement set](/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets)
- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md)
- [Nested app authentication and Outlook legacy tokens deprecation FAQ](faq-nested-app-auth-outlook-legacy-tokens.md)
