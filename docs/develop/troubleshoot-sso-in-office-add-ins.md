---
title: Troubleshoot error messages for single sign-on (SSO)
description: 
ms.date: 11/20/2017 
---

# Troubleshoot error messages for single sign-on (SSO) (preview)

This article provides some guidance about how to troubleshoot problems with single sign-on (SSO) in Office Add-ins, and how to make your SSO-enabled add-in robustly handle special conditions or errors.

## Debugging tools

We strongly recommend that you use a tool that can intercept and display the HTTP Requests from, and Responses to, your add-in's web service when you are developing. Two of the most popular are: 

- [Fiddler](http://www.telerik.com/fiddler): Free ([Documentation](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): Free for 30 days. ([Documenation](https://www.charlesproxy.com/documentation/))

When developing your service API, you may also want to try:

- [Postman](http://www.getpostman.com/postman): Free ([Documentation](https://www.getpostman.com/docs/))

## Causes and handling of errors from getAccessTokenAsync

### 13000

The [getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync) API is not supported by the add-in or the Office version. 

- The version of Office does not support SSO. The required version is Office 2016, Version 1710, build 8629.nnnn or later (the Office 365 subscription version, sometimes called “Click to Run”). You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1). 
- The add-in manifest is missing the proper [WebApplicationInfo](https://dev.office.com/reference/add-ins/manifest/webapplicationinfo) section.

### 13001

The user is not signed into Office. Your code should recall the `getAccessTokenAsync` method and pass the option `forceAddAccount: true` in the [options](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync#parameters) parameter. 

### 13002

The user aborted sign in or consent. 
- If your add-in provides functions that don't require the user to be signed in (or to have granted consent), then your code should catch this error and allow the add-in to stay running.
- If the add-in requires a signed-in user who has granted consent, your code should ask the user to repeat the operation, but not more than once. 

### 13003

User Type not supported. The user isn't signed into Office with a valid Microsoft Account or Work or School account. This may happen if Office runs with an on-premises domain account, for example. Your code should ask the user to sign in to Office.

### 13004

Invalid Resource. The add-in manifest hasn’t been configured correctly. Update the manifest. For more information, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md).

### 13005

Invalid Grant. This usually means that Office has not been pre-authorized to the add-in's web service. For more information, see [Create the service application](sso-in-office-add-ins.md#create-the-service-application) and [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) or [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS). This also may happen if the user has not granted your service application permissions to his or her `profile`.

### 13006

Client Error. Your code should suggest that the user sign out and restart Office.

### 13007

The Office host was unable to get an access token to the add-in's web service.
- Be sure that your add-in registration and add-in manifest specify the `openid` and `profile` permissions. For more information, see [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) or [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS), and [Configure the add-in](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) or [Configure the add-in](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS).
- Your code could suggest that the user retry the operation later.

### 13008

The user triggered an operation that calls `getAccessTokenAsync` before a previous call of `getAccessTokenAsync` completed. Your code should ask the user to repeat the operation after the previous operation has completed.

## Errors on the server-side from Azure Active Directory

### Conditional access / Multifactor authentication errors
 
In certain configurations of identity in AAD and Office 365, it is possible for some resources that are accessible with Microsoft Graph to require multifactor authentication (MFA), even when the user's Office 365 tenancy does not. When AAD receives a request for a token to the MFA-protected resource, via the on-behalf-of flow, it returns to your add-in's web service a JSON message that contains a `claims` property. The claims property has information about what further authentication factors are needed. 

Your server-side code should test for this message and relay the claims value to your client-side code. You need this information in the client because Office handles authentication for SSO add-ins. The message to the client can be either an error (such as `500 Server Error` or `401 Unauthorized`) or in the body of a success response (such as `200 OK`). In either case, the (failure or success) callback of your code's client-side AJAX call to your add-in's web API should test for this response. If the claims value has been relayed, your code should recall `getAccessTokenAsync` and pass the option `authChallenge: CLAIMS-STRING-HERE` in the [options](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync#parameters) parameter. When AAD sees this string, it prompts the user for the additional factor(s) and then returns a new access token which will be accepted in the on-behalf-of flow.

We have some samples to illustrate this MFA handling: 

- [Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO): The MSAL library that this sample uses exposes the MFA message from AAD as an exception. The code relays this to the client as a `500 Server Error` response. In the client-side script, the `fail` callback of the AJAX call recalls `getAccessTokenAsync` with the `authChallenge` option. See especially the files [ValuesController.cs](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Controllers/ValuesController.cs) and [Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js).
- [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO): The MFA message from AAD is sent to the client as a success response. In the client-side script, the `done` callback of the AJAX call recalls `getAccessTokenAsync` with the `authChallenge` option. See especially the files [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) and [program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).

### Consent missing errors

If AAD has no record that consent (to the Microsoft Graph resource) was granted to the add-in by the user (or tenant administrator), AAD will send an error message to your web service. Your code must tell the client (in the body of a `403 Forbidden` response, for example) to recall `getAccessTokenAsync` with the `forceConsent: true` option.

### Invalid or missing scope (permission) errors

- Your server-side code should send a `403 Forbidden` response to the client which should present a friendly message to the user. If possible, log the error to the console or record it in a log.
- Be sure your add-in manifest [Scopes](https://dev.office.com/reference/add-ins/manifest/scopes)  section specifies all needed permissions. And be sure your registration of the add-in's web service specifies the same permissions. Check for spelling mistakes too. For more information, see [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) or [Register the add-in with Azure AD v2.0 endpoint](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS), and [Configure the add-in](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) or [Configure the add-in](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS).

### Expired or invalid token errors when calling Microsoft Graph

Some authentication and authorization libraries, including MSAL, prevent expired token errors by using a cached refresh token whenever necessary. You can also code your own token caching system. For a sample that does this, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especially the file [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

But if you get an expired token or invalid token error, your code must tell the client (in the body of a `401 Unauthorized` response, for example) to recall `getAccessTokenAsync` and repeat the call to the endpoint of your add-in's web API, which will repeat the on-behalf-of flow to obtain a new token for Microsoft Graph. 

### Invalid token error when calling Microsoft Graph

Handle this error the same as an expired token error. See previous section.

### Invalid audience error

Your server-side code should send a `403 Forbidden` response to the client which should present a friendly message to the user and possibly also log the error to the console or record it in a log.

For more on adding multitenant support for token validation, see the [Azure Multitenant Sample](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
