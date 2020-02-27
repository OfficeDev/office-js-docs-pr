---
title: Authenticate a user with a single-sign-on token
description: Learn about using the single-sign-on token provided by an Outlook add-in to implement SSO with your service.
ms.date: 11/19/2019
localization_priority: Normal
---

# Authenticate a user with a single-sign-on token in an Outlook add-in (preview)

Single sign-on (SSO) provides a seamless way for your add-in to authenticate users (and optionally to obtain access tokens to call the [Microsoft Graph API](/graph/overview)).

Using this method, your add-in can obtain an access token scoped to your server back-end API. The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API. Optionally, you can also have your server-side code:

- Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API
- Use the identity information in the token to establish the user's identity and authenticate to your own back-end services

For an overview of SSO in Office Add-ins, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) and [Authorize to Microsoft Graph in your Office Add-in](../develop/authorize-to-microsoft-graph.md).

> [!NOTE]
> To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.

## Enable modern authentication in your Office 365 tenancy

To use SSO with an Outlook add-in, you must enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## Register your add-in

To use SSO, your Outlook add-in will need to have a server-side web API that is registered with Azure Active Directory (AAD) v2.0. For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](../develop/register-sso-add-in-aad-v2.md).

### Provide consent when sideloading an add-in

When an add-in that uses SSO is acquired from AppSource, the store UI handles prompting the user for consent to the requested Graph permissions. However, when you are developing an add-in, you have to provide consent in advance. For more information, see [Grant administrator consent to the add-in](../develop/grant-admin-consent-to-an-add-in.md)

## Update the add-in manifest

The next step to enable SSO in the add-in is to add a `WebApplicationInfo` element at the end of the `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md) element. For more information, see [Configure the add-in](../develop/sso-in-office-add-ins.md#configure-the-add-in).

## Get the SSO token

The add-in gets an SSO token with client-side script. For more information, see [Add client-side code](../develop/sso-in-office-add-ins.md#add-client-side-code).

## Use the SSO token at the back-end

In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there. For details on what your server-side could and should do, see [Add server-side code](../develop/sso-in-office-add-ins.md#add-server-side-code).

> [!IMPORTANT]
> When using the SSO token as an identity in an *Outlook* add-in, we recommend that you also [use the Exchange identity token](authenticate-a-user-with-an-identity-token.md) as an alternate identity. Users of your add-in may use multiple clients, and some may not support providing an SSO token. By using the Exchange identity token as an alternate, you can avoid having to prompt these users for credentials multiple times. For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](implement-sso-in-outlook-add-in.md).

## See also

- For a sample Outlook add-in that uses the SSO token to access the Microsoft Graph API, see [AttachmentsDemo Sample Add-in](https://github.com/OfficeDev/outlook-add-in-attachments-demo).
- [SSO API reference](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [IdentityAPI requirement set](../reference/requirement-sets/identity-api-requirement-sets.md)
