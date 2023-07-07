---
title: Enable single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based Outlook add-in
description: Learn how to enable SSO or CORS when working in an event-based activation add-in.
ms.date: 07/06/2023
ms.localizationpriority: medium
---

# Enable single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based Outlook add-in

When an Outlook add-in uses event-based activation, the events run in a separate [runtime](../testing/runtimes.md). To enable single sign-on (SSO) in your event-based add-in or allow it to request external data through cross-origin resource sharing (CORS), you must configure a well-known URI. Through this resource, Office will be able to identify the add-ins, including their JavaScript files, that support SSO or CORS requests.

> [!NOTE]
> The steps in this article only apply when running your Outlook add-in on Windows. This is because Outlook on Windows uses a JavaScript file, while Outlook on Mac and on the web use an HTML file that references the same JavaScript file. To learn more about event-based activation in Outlook add-ins, see [Configure your Outlook add-in for event-based activation](autolaunch.md).

## List allowed add-ins in a well-known URI

To list which add-ins are allowed to work with SSO or CORS, create a JSON file that identifies each JavaScript file for each add-in. Then, host that JSON file at a well-known URI. A well-known URI allows the specification of all hosted JS files that are authorized to obtain tokens for the current web origin. This ensures that the owner of the origin has full control over which hosted JavaScript files are meant to be used in an add-in and which ones are not, preventing any security vulnerabilities around impersonation, for example.

The following example shows how to enable SSO or CORS for two add-ins (a main version and beta version). You can list as many add-ins as necessary depending on how many you provide from your web server.

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

Host the JSON file under a location named `.well-known` in the URI at the root of the origin. For example, if the origin is `https://addin.contoso.com:8000/`, then the well-known URI is `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json`.

The origin refers to a pattern of scheme + subdomain + domain + port. The name of the location **must** be `.well-known`, and the name of the resource file **must** be `microsoft-officeaddins-allowed.json`. This file must contain a JSON object with an attribute named `allowed` whose value is an array of all JavaScript files authorized for SSO for their respective add-ins.

After you configure the well-known URI, if your add-in implements SSO, you can then call the [getAccessToken() API](/javascript/api/office-runtime/officeruntime.auth) to get an access token with the user's identity.

> [!IMPORTANT]
> While `OfficeRuntime.auth.getAccessToken` and `Office.auth.getAccessToken` perform the same functionality of retrieving an access token, we recommend calling `OfficeRuntime.auth.getAccessToken` in your event-based add-in. This API is supported in all Outlook client versions that support event-based activation and SSO. On the other hand, `Office.auth.getAccessToken` is only supported in Outlook on Windows starting from Version 2111 (Build 14701.20000).

## See also

- [Authenticate a user with a single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md)
- [Configure your Outlook add-in for event-based activation](autolaunch.md)
