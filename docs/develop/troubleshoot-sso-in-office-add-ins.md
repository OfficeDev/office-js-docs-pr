---
title: Troubleshoot error messages for legacy Office single sign-on (SSO)
description: Guidance about how to troubleshoot problems with single sign-on (SSO) in Office Add-ins, and handle special conditions or errors.
ms.date: 01/12/2026
ms.localizationpriority: medium
---

# Troubleshoot error messages for legacy Office single sign-on (SSO)

This article provides some guidance about how to troubleshoot problems with legacy Office single sign-on (SSO) in Office Add-ins, and how to make your SSO-enabled add-in robustly handle special conditions or errors.

> [!NOTE]
> This article describes legacy Office single sign-on (SSO). For a modern authentication experience with support across a wider range of platforms, use the Microsoft Authentication Library (MSAL) with nested app authentication (NAA). For more information, see [Enable single sign-on in an Office Add-in with nested app authentication](enable-nested-app-authentication-in-your-add-in.md).

> [!NOTE]
> The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
> If you're working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Enable or disable modern authentication for Outlook in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online).

## Debugging tools

We strongly recommend that you use a tool that can intercept and display the HTTP Requests from, and Responses to, your add-in's web service when you are developing. Some of the most popular are:

- [Fiddler](https://www.telerik.com/fiddler): Free for 10 days ([Documentation](https://docs.telerik.com/fiddler-everywhere/introduction))
- [Charles](https://www.charlesproxy.com): Free for 30 days. ([Documentation](https://www.charlesproxy.com/documentation))
- [Requestly](https://requestly.com/downloads): Free basic service. ([Documentation](https://developers.requestly.com))

## Causes and handling of errors from getAccessToken

For examples of the error handling described in this section, see:

- [HomeES6.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNETCore-WebAPI/wwwroot/js/HomeES6.js)
- [ssoAuthES6.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/public/javascripts/ssoAuthES6.js)

### 13000

The [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) API isn't supported by the add-in or the Office version.

- The version of Office does not support SSO. The required version is Microsoft 365 subscription, in any monthly channel.
- The add-in manifest is missing the proper [WebApplicationInfo](/javascript/api/manifest/webapplicationinfo) section.

Your add-in should respond to this error by falling back to an alternate system of user authentication. For more information, see [Requirements and Best Practices](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

### 13001

The user isn't signed into Office. In most scenarios, you should prevent this error from ever being seen by passing the option `allowSignInPrompt: true` in the `AuthOptions` parameter.

But there may be exceptions. For example, you want the add-in to open with features that require a logged in user; but *only if* the user is *already* logged into Office. If the user isn't logged in, you want the add-in to open with an alternate set of features that do not require that the user is signed in. In this case, logic which runs when the add-in launches calls `getAccessToken` without `allowSignInPrompt: true`. Use the 13001 error as the flag to tell the add-in to present the alternate set of features.

Another option is to respond to 13001 by falling back to an alternate system of user authentication. This will sign the user into Microsoft Entra ID, but not sign the user into Office.

This error doesn't typically occur in Office on the web. If the user's cookie expires, Office on the web returns [error 13006](#13006). However, if a user accesses Outlook on the web from Firefox with Enhanced Tracking Protection turned on, they'll encounter error 13001.

### 13002

The user aborted sign in or consent; for example, by choosing **Cancel** on the consent dialog.

- If your add-in provides functions that don't require the user to be signed in (or to have granted consent), then your code should catch this error and allow the add-in to stay running.
- If the add-in requires a signed-in user who has granted consent, your code should have a sign-in button appear.

### 13003

User Type not supported. The user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or work account. This may happen if Office runs with an on-premises domain account, for example. Your code should fall back to an alternate system of user authentication. In Outlook, this error may also occur if [modern authentication is disabled](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online) for the user's tenant in Exchange Online. For more information, see [Requirements and Best Practices](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

### 13004

Invalid Resource. (This error should only be seen in development.) The add-in manifest hasn't been configured correctly. Update the manifest. For more information, see [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md). The most common problem is that the `<Resource>` element (in the `<WebApplicationInfo>` element) has a domain that does not match the domain of the add-in. Although the protocol part of the Resource value should be "api" not "https"; all other parts of the domain name (including port, if any) should be the same as for the add-in.

### 13005

Invalid Grant. This usually means that Office has not been pre-authorized to the add-in's web service. For more information, see [Create the service application](sso-in-office-add-ins.md#register-your-add-in-with-the-microsoft-identity-platform) and [Register an Office Add-in that uses single sign-on (SSO) with the Microsoft identity platform](register-sso-add-in-aad-v2.md). This also may happen if the user has not granted your service application permissions to their `profile`, or has revoked consent. Your code should fall back to an alternate system of user authentication.

Another possible cause, during development, is that your add-in using Internet Explorer, and you are using a self-signed certificate. (To determine which browser or webview is being used by the add-in, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).)

### 13006

Client Error. This error is only seen in **Office on the web**. Your code should suggest that the user sign out and then restart the Office browser session.

### 13007

The Office application was unable to get an access token to the add-in's web service.

- If this error occurs during development, be sure that your add-in registration and add-in manifest specify the `profile` permission (and the `openid` permission, if you are using MSAL.NET). For more information, see [Register an Office Add-in that uses single sign-on (SSO) with the Microsoft identity platform](register-sso-add-in-aad-v2.md).
- In production, an account mismatch could cause this error. For example, if the user attempts to sign in with a personal Microsoft account (MSA) when a Work or school account was expected. For these cases, your code should fall back to an alternate system of user authentication. For more information on account types, see [Identity and account types for single- and multi-tenant apps](/security/zero-trust/develop/identity-supported-account-types).
- Make sure your application is enabled for users to sign-in for your organization.
    1. Sign in to the [Microsoft Azure portal](https://portal.azure.com/).
    1. Go to your add-in's app registration.
    1. On the **Overview** page,  select **Managed application in local directory**.
        :::image type="content" source="../images/azure-portal-managed-application.png" alt-text="The Managed application in local directory option in the App Registration Overview window.":::
    1. Select **Manage** > **Properties**, and ensure that the value of **Enabled for users to sign-in?** is **Yes**.
        :::image type="content" source="../images/azure-portal-enable-sign-in.png" alt-text="The option to allow users in the organization to sign-in to an application in the Properties window.":::

### 13008

The user triggered an operation that calls `getAccessToken` before a previous call of `getAccessToken` completed. This error is only seen on **Office on the web**. Your code should ask the user to repeat the operation after the previous operation has completed.

### 13010

The user is running the add-in in Office on Microsoft Edge. The user's Microsoft 365 domain, and the `login.microsoftonline.com` domain, are in a different security zones in the browser settings. This error is only seen on **Office on the web**. If this error is returned, the user will have already seen an error explaining this and linking to a page about how to change the zone configuration. If your add-in provides functions that don't require the user to be signed in, then your code should catch this error and allow the add-in to stay running.

### 13012

There are several possible causes.

- The add-in is running on a platform that does not support the `getAccessToken` API. For example, it isn't supported on iPad. See also [Identity API requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
- The Office document was opened from the **Files** tab of a Teams channel using the **Edit in Teams** option on the **Open** dropdown menu. The `getAccessToken` API isn't supported in this scenario.
- The `forMSGraphAccess` option was passed in the call to `getAccessToken` and the user obtained the add-in from Microsoft Marketplace. In this scenario, the tenant admin has not granted consent to the add-in for the Microsoft Graph scopes (permissions) that it needs. Recalling `getAccessToken` with the `allowConsentPrompt` will not solve the problem because Office is allowed to prompt the user for consent to only the Microsoft Entra ID `profile` scope.

Your code should fall back to an alternate system of user authentication.

In development, the add-in is sideloaded in Outlook and the `forMSGraphAccess` option was passed in the call to `getAccessToken`.

### 13013

The `getAccessToken` was called too many times in a short amount of time, so Office throttled the most recent call. This is usually caused by an infinite loop of calls to the method. There are scenarios when recalling the method is advisable. However, your code should use a counter or flag variable to ensure that the method isn't recalled repeatedly. If the same "retry" code path is running again, the code should fall back to an alternate system of user authentication. For a code example, see how the `retryGetAccessToken` variable is used in [HomeES6.js](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNETCore-WebAPI/wwwroot/js/HomeES6.js) or [ssoAuthES6.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/public/javascripts/ssoAuthES6.js).

### 50001

This error (which isn't specific to `getAccessToken`) may indicate that the browser has cached an old copy of the office.js files. When you're developing, clear the browser's cache. Another possibility is that the version of Office isn't recent enough to support SSO. On Windows, the minimum version is Version 1911 (Build 12215.20006). On Mac, it's Version 16.32 (19102902).

In a production add-in, the add-in should respond to this error by falling back to an alternate system of user authentication. For more information, see [Requirements and Best Practices](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

## Errors on the server-side from Microsoft Entra ID

For samples of the error-handling described in this section, see:

- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)

### Conditional access / Multifactor authentication errors

In certain configurations of identity in Microsoft Entra ID and Microsoft 365, it is possible for some resources that are accessible with Microsoft Graph to require multifactor authentication (MFA), even when the user's Microsoft 365 tenancy does not. When Microsoft Entra ID receives a request for a token to the MFA-protected resource, via the on-behalf-of flow, it returns to your add-in's web service a JSON message that contains a `claims` property. The claims property has information about what further authentication factors are needed.

Your code should test for this `claims` property. Depending on your add-in's architecture, you may test for it on the client-side, or you may test for it on the server-side and relay it to the client. You need this information in the client because Office handles authentication for SSO add-ins. If you relay it from the server-side, the message to the client can be either an error (such as `500 Server Error` or `401 Unauthorized`) or in the body of a success response (such as `200 OK`). In either case, the (failure or success) callback of your code's client-side AJAX call to your add-in's web API should test for this response.

Regardless of your architecture, if the claims value has been sent from Microsoft Entra ID, your code should recall `getAccessToken` and pass the option `authChallenge: CLAIMS-STRING-HERE` in the `options` parameter. When Microsoft Entra ID sees this string, it prompts the user for the additional factors and then returns a new access token which will be accepted in the on-behalf-of flow.

### Consent missing errors

If Microsoft Entra ID has no record that consent (to the Microsoft Graph resource) was granted to the add-in by the user (or tenant administrator), Microsoft Entra ID will send an error message to your web service. Your code must tell the client (in the body of a `403 Forbidden` response, for example).

If the add-in needs Microsoft Graph scopes that can only be consented to by an admin, your code should throw an error. If the only scopes that are needed can be consented to by the user, then your code should fall back to an alternate system of user authentication.

### Invalid or missing scope (permission) errors

This kind of error should only be seen in development.

- Your server-side code should send a `403 Forbidden` response to the client which should log the error to the console or record it in a log.
- Be sure your add-in manifest [Scopes](/javascript/api/manifest/scopes) section specifies all needed permissions. And be sure your registration of the add-in's web service specifies the same permissions. Check for spelling mistakes too. For more information, see [Register an Office Add-in that uses single sign-on (SSO) with the Microsoft identity platform](register-sso-add-in-aad-v2.md).

### Invalid audience error in the access token for Microsoft Graph

Your server-side code should send a `403 Forbidden` response to the client which should present a friendly message to the user and possibly also log the error to the console or record it in a log.
