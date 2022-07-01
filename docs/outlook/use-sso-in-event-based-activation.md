---
title: Enable single sign-on (SSO) in Outlook add-ins that use event-based activation
description: Learn how to enable SSO when working in an event-based activation add-in.
ms.date: 06/17/2022
ms.localizationpriority: medium
---

# Enable single sign-on (SSO) in Outlook add-ins that use event-based activation

When an Outlook add-in uses event-based activation, the events run in a separate JavaScript runtime. After completing the steps in [Authenticate a user with a single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md), follow the additional steps described in this article to enable SSO for your event handling code. Once you enable SSO you can call the [getAccessToken() API](/javascript/api/office-runtime/officeruntime.auth) to get an access token with the user's identity.

> [!IMPORTANT]
> While `OfficeRuntime.auth.getAccessToken` and `Office.auth.getAccessToken` perform the same functionality of retrieving an access token, we recommend calling `OfficeRuntime.auth.getAccessToken` in your event-based add-in. This API is supported in all Outlook client versions that support event-based activation and SSO. On the other hand, `Office.auth.getAccessToken` is only supported in Outlook on Windows starting from version 2111 (build 14701.20000).

For Outlook on Windows, in the manifest for your Outlook add-in, you identify a single JavaScript file to load for event-based activation. You also need to specify to Office that this file is allowed to support SSO. You do this by creating a list of all add-ins, and their JavaScript files, to provide to Office through a well-known URI.

> [!NOTE]
> The steps in this article only apply when running your Outlook add-in on Windows. This is because Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.

## List allowed add-ins with a well-known URI

To list which add-ins are allowed to work with SSO, create a JSON file that identifies each JavaScript file for each add-in. Then host that JSON file at a well-known URI. A well-known URI allows the specification of all hosted JS files that are authorized to obtain tokens for the current web origin. This ensures that the owner of the origin has full control over which hosted JS files are meant to be used in an add-in and which ones are not, preventing any security vulnerabilities around impersonation, for example.

The following example shows how to enable SSO for two add-ins (a main version and beta version). You can list as many add-ins as necessary depending on how many you provide from your web server.

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

## See also

- [Authenticate a user with a single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md)
- [Configure your Outlook add-in for event-based activation](autolaunch.md)
