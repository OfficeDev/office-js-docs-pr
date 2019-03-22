---
title: Troubleshoot error messages for single sign-on (SSO)
description: ''
ms.date: 03/21/2019
localization_priority: Priority
---

# Troubleshoot error messages for single sign-on (SSO) (preview)

This article provides some guidance about how to troubleshoot problems with single sign-on (SSO) in Office Add-ins, and how to make your SSO-enabled add-in robustly handle special conditions or errors.

> [!NOTE]
> The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).
> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## Debugging tools

We strongly recommend that you use a tool that can intercept and display the HTTP Requests from, and Responses to, your add-in's web service when you are developing. Two of the most popular are:

- [Fiddler](https://www.telerik.com/fiddler): Free ([Documentation](https://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): Free for 30 days. ([Documentation](https://www.charlesproxy.com/documentation/))

When developing your service API, you may also want to try:

- [Postman](https://www.getpostman.com/postman): Free ([Documentation](https://www.getpostman.com/docs/))

## Causes and handling of errors from getAccessTokenAsync

For examples of the error handling described in this section, see:
- [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> Besides the suggestions made in this section, an Outlook add-in has an additional way to respond to any 13*nnn* error. For details, see [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in) and [AttachmentsDemo Sample Add-in](https://github.com/OfficeDev/outlook-add-in-attachments-demo).

### 13000

The [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) API is not supported by the add-in or the Office version.

- The version of Office does not support SSO. The required version is Office 365 (the subscription version of Office), Version 1710, build 8629.nnnn or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).
- The add-in manifest is missing the proper [WebApplicationInfo](/office/dev/add-ins/reference/manifest/webapplicationinfo) section.

Your add-in should respond to this error by falling back to an alternate system of user authentication. For more information, see [Requirements and Best Practices](/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices).

### 13001

The user is not signed into Office. Your code should recall the `getAccessTokenAsync` method and pass the option `forceAddAccount: true` in the [options](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) parameter. But don't do this more than once. The user may have decided not to sign-in.

This error is never seen in Office Online. If the user's cookie expires, Office Online returns error 13006.

### 13002

The user aborted sign in or consent; for example, by choosing **Cancel** on the consent dialog.

- If your add-in provides functions that don't require the user to be signed in (or to have granted consent), then your code should catch this error and allow the add-in to stay running.
- If the add-in requires a signed-in user who has granted consent, your code should ask the user to repeat the operation, but not more than once.

### 13003

User Type not supported. The user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account. This may happen if Office runs with an on-premises domain account, for example. Your code should either ask the user to sign in to Office or fall back to an alternate system of user authentication. For more information, see [Requirements and Best Practices](/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices).

### 13004

Invalid Resource. The add-in manifest hasn’t been configured correctly. Update the manifest. For more information, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). The most common problem is that the **Resource** element (in the **WebApplicationInfo** element) has a domain that does not match the domain of the add-in. Although the protocol part of the Resource value should be "api" not "https"; all other parts of the domain name (including port, if any) should be the same as for the add-in.

### 13005

Invalid Grant. This usually means that Office has not been pre-authorized to the add-in's web service. For more information, see [Create the service application](sso-in-office-add-ins.md#create-the-service-application) and [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) or [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS). This also may happen if the user has not granted your service application permissions to their `profile`.

### 13006

Client Error. Your code should suggest that the user sign out and restart Office, or restart the Office Online session.

### 13007

The Office host was unable to get an access token to the add-in's web service.

- If this error occurs during development, be sure that your add-in registration and add-in manifest specify the `openid` and `profile` permissions. For more information, see [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) or [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS), and [Configure the add-in](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) or [Configure the add-in](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS).
- In production, there are several things that can cause this error. Some of them are:
    - The user has revoked consent, after previously granting it. Your code should recall the `getAccessTokenAsync` method with the option `forceConsent: true`, but no more than once.
    - The user is has an Microsoft Account (MSA) identity. Some situations that would cause one of the other 13nnn errors with a Work or School account, will cause a 13007 when a MSA is used.

  For all of these cases, if you have already tried the `forceConsent` option once, then your code could suggest that the user retry the operation later.

### 13008

The user triggered an operation that calls `getAccessTokenAsync` before a previous call of `getAccessTokenAsync` completed. Your code should ask the user to repeat the operation after the previous operation has completed.

### 13009

The add-in called the `getAccessTokenAsync` method with the option `forceConsent: true`, but the add-in's manifest is deployed to a type of catalog that does not support forcing consent. Your code should recall the `getAccessTokenAsync` method and pass the option `forceConsent: false` in the [options](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) parameter. However, the call of  `getAccessTokenAsync`  with `forceConsent: true` might itself have been an automatic response to a failed call of `getAccessTokenAsync` with `forceConsent: false`, so your code should keep track of whether `getAccessTokenAsync` with `forceConsent: false` has already been called. If it has, your code should either tell the user to sign out of Office and sign-in again or it should fall back to an alternate system of user authentication. For more information, see [Requirements and Best Practices](/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices).

> [!NOTE]
> Microsoft will not necessarily impose this restriction on any types of add-in catalogs. If it doesn't, then this error will never be seen.

### 13010

The user is running the add-in on Office Online and is using Edge or Internet Explorer. The user’s Office 365 domain, and the login.microsoftonline.com domain, are in a different security zones in the browser settings. If this error is returned, the user will have already seen an error explaining this and linking to a page about how to change the zone configuration. If your add-in provides functions that don't require the user to be signed in, then your code should catch this error and allow the add-in to stay running.

### 13012

The add-in is running on a platform that does not support the `getAccessTokenAsync` API. For example, it is not supported on iPad. See also [Identity API Requirement Sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).

### 50001

This error (which is not specific to `getAccessTokenAsync`) may indicate that the browser has cached an old copy of the office.js files. When you are developing, clear the browser's cache. Another possibility is that the version of Office is not recent enough to support SSO. See [Prerequisites](create-sso-office-add-ins-aspnet.md#prerequisites).

In a production add-in, the add-in should respond to this error by falling back to an alternate system of user authentication. For more information, see [Requirements and Best Practices](/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices).


## Errors on the server-side from Azure Active Directory

For samples of the error-handling described in this section, see:
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### Conditional access / Multifactor authentication errors

In certain configurations of identity in AAD and Office 365, it is possible for some resources that are accessible with Microsoft Graph to require multifactor authentication (MFA), even when the user's Office 365 tenancy does not. When AAD receives a request for a token to the MFA-protected resource, via the on-behalf-of flow, it returns to your add-in's web service a JSON message that contains a `claims` property. The claims property has information about what further authentication factors are needed.

Your server-side code should test for this message and relay the claims value to your client-side code. You need this information in the client because Office handles authentication for SSO add-ins. The message to the client can be either an error (such as `500 Server Error` or `401 Unauthorized`) or in the body of a success response (such as `200 OK`). In either case, the (failure or success) callback of your code's client-side AJAX call to your add-in's web API should test for this response. If the claims value has been relayed, your code should recall `getAccessTokenAsync` and pass the option `authChallenge: CLAIMS-STRING-HERE` in the [options](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) parameter. When AAD sees this string, it prompts the user for the additional factor(s) and then returns a new access token which will be accepted in the on-behalf-of flow.

### Consent missing errors

If AAD has no record that consent (to the Microsoft Graph resource) was granted to the add-in by the user (or tenant administrator), AAD will send an error message to your web service. Your code must tell the client (in the body of a `403 Forbidden` response, for example) to recall `getAccessTokenAsync` with the `forceConsent: true` option.

### Invalid or missing scope (permission) errors

- Your server-side code should send a `403 Forbidden` response to the client which should present a friendly message to the user. If possible, log the error to the console or record it in a log.
- Be sure your add-in manifest [Scopes](/office/dev/add-ins/reference/manifest/scopes) section specifies all needed permissions. And be sure your registration of the add-in's web service specifies the same permissions. Check for spelling mistakes too. For more information, see [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) or [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS), and [Configure the add-in](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) or [Configure the add-in](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS).

### Expired or invalid token errors when calling Microsoft Graph

Some authentication and authorization libraries, including MSAL, prevent expired token errors by using a cached refresh token whenever necessary. You can also code your own token caching system. For a sample that does this, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especially the file [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

But if you get an expired token or invalid token error, your code must tell the client (in the body of a `401 Unauthorized` response, for example) to recall `getAccessTokenAsync` and repeat the call to the endpoint of your add-in's web API, which will repeat the on-behalf-of flow to obtain a new token for Microsoft Graph.

### Invalid token error when calling Microsoft Graph

Handle this error the same as an expired token error. See previous section.

### Invalid audience error

Your server-side code should send a `403 Forbidden` response to the client which should present a friendly message to the user and possibly also log the error to the console or record it in a log.

For more on adding multitenant support for token validation, see the [Azure Multitenant Sample](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
