---
title: Authenticate a user with a single-sign-on token
description: Learn about using the single-sign-on token provided by an Outlook add-in to implement SSO with your service.
ms.date: 06/24/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Authenticate a user with a single-sign-on token in an Outlook add-in

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

Single sign-on (SSO) provides a seamless way for your add-in to authenticate users (and optionally to obtain access tokens to call the [Microsoft Graph API](/graph/overview)).

Using this method, your add-in can obtain an access token scoped to your server back-end API. The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API. Optionally, you can also have your server-side code.

- Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API
- Use the identity information in the token to establish the user's identity and authenticate to your own back-end services

For an overview of SSO in Office Add-ins, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) and [Authorize to Microsoft Graph in your Office Add-in](../develop/authorize-to-microsoft-graph.md).

## Enable modern authentication in your Microsoft 365 tenancy

To use SSO with an Outlook add-in, you must enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Enable or disable modern authentication for Outlook in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online).

## Register your add-in

To use SSO, your Outlook add-in will need to have a server-side web API that is registered with Microsoft Entra ID. For more information, see [Register an Office Add-in that uses SSO with the Microsoft identity platform](../develop/register-sso-add-in-aad-v2.md).

### Provide consent when sideloading an add-in

When you are developing an add-in, you will have to provide consent in advance. For more information, see [Admin consent](/entra/identity-platform/quickstart-configure-app-access-web-apis#more-on-api-permissions-and-admin-consent).

## Update the add-in manifest

The next step to enable SSO in the add-in is to add some information to the manifest from the add-in's Microsoft identity platform registration. The markup varies depending on the type of manifest.

- **Add-in only manifest**: Add a `WebApplicationInfo` element at the end of the `VersionOverridesV1_1` [VersionOverrides](/javascript/api/manifest/versionoverrides) element. Then, add its required child elements. For detailed information about the markup, see [Configure the add-in](../develop/sso-in-office-add-ins.md#configure-the-add-in).
- **Unified manifest for Microsoft 365**: Add a [`"webApplicationInfo"`](/microsoft-365/extensibility/schema/root#webApplicationInfo-property) property to the root `{ ... }` object in the manifest. Give this object a child `"id"` property set to the application ID of the add-in's web app as it was generated in the Azure portal when you registered the add-in. (See the section [Register your add-in](#register-your-add-in) earlier in this article.) Also give it a child `"resource"` property that is set to the same **Application ID URI** that you set when you registered the add-in. This URI should have the form `api://<fully-qualified-domain-name>/<application-id>`. The following is an example.

   ```json
   "webApplicationInfo": {
        "id": "a661fed9-f33d-4e95-b6cf-624a34a2f51d",
        "resource": "api://addin.contoso.com/a661fed9-f33d-4e95-b6cf-624a34a2f51d"
    },
   ```

## Get the SSO token

The add-in gets an SSO token with client-side script. For more information, see [Add client-side code](../develop/sso-in-office-add-ins.md#add-client-side-code).

## Use the SSO token at the back-end

In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there. For details on what your server-side could and should do, see [Add server-side code](../develop/sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).

## SSO for event-based activation or integrated spam reporting

There are additional steps to take if your add-in uses event-based activation or integrated spam reporting. For more information, see [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Outlook add-in](../develop/use-sso-in-event-based-activation.md).

## See also

- [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1))
- For a sample Outlook add-in that uses the SSO token to access the Microsoft Graph API, see [Outlook Add-in SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).
- [SSO API reference](/javascript/api/office/office.auth#office-office-auth-getaccesstoken-member(1))
- [IdentityAPI requirement set](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Outlook add-in](../develop/use-sso-in-event-based-activation.md)
