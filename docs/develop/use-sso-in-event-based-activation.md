---
title: Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Office Add-in
description: Learn how to use SSO or CORS in an add-in that implements event-based activation or integrated spam reporting.
ms.date: 07/08/2025
ms.localizationpriority: medium
---

# Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Office Add-in

When an add-in implements event-based activation or integrated spam reporting, the events run in a separate [runtime](../testing/runtimes.md). To configure single sign-on (SSO) or request external data through cross-origin resource sharing (CORS) in these add-ins, you must configure a well-known URI. Through this resource, Office will be able to identify the add-ins, including their JavaScript files, that support SSO or CORS requests.

> [!NOTE]
> The steps in this article only apply to add-ins that run on Excel, PowerPoint, or Word on Windows, or classic Outlook on Windows. This is because they use a JavaScript file, while these applications on Mac, on the web, and [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) use an HTML file that references the same JavaScript file. To learn more, see [Activate add-ins with events](../develop/event-based-activation.md) and [Implement an integrated spam-reporting add-in](../outlook/spam-reporting.md).

## List allowed add-ins in a well-known URI

To list which add-ins are allowed to work with SSO or CORS, create a JSON file that identifies each JavaScript file for each add-in. Then, host that JSON file at a well-known URI. A well-known URI allows the specification of all hosted JS files that are authorized to obtain tokens for the current web origin. This ensures that the owner of the origin has full control over which hosted JavaScript files are meant to be used in an add-in and which ones are not, preventing any security vulnerabilities around impersonation, for example.

The following example shows how to configure SSO or CORS for two add-ins (a main version and beta version). You can list as many add-ins as necessary depending on how many you provide from your web server.

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

Host the JSON file under a location named `.well-known` in the URI at the root of the origin. For example, if the origin is `https://addin.contoso.com:8000/`, then the well-known URI is `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json`. For clarification, this file is to be hosted in your Office Web Add-in, not the web server that you're attempting to make a CORS request to. See the [Outlook-Add-in-SSO-events](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/auth/Outlook-Add-in-SSO-events/public/.well-known/microsoft-officeaddins-allowed.json) sample for an example using the recommended location.

The origin refers to a pattern of scheme + subdomain + domain + port. The name of the location **must** be `.well-known`, and the name of the resource file **must** be `microsoft-officeaddins-allowed.json`. This file must contain a JSON object with an attribute named `allowed` whose value is an array of all JavaScript files authorized for SSO for their respective add-ins.

After you configure the well-known URI, if your add-in implements SSO, you can then call the [getAccessToken() API](/javascript/api/office-runtime/officeruntime.auth) to get an access token with the user's identity.

> [!IMPORTANT]
> While `OfficeRuntime.auth.getAccessToken` and `Office.auth.getAccessToken` perform the same functionality of retrieving an access token, we recommend calling `OfficeRuntime.auth.getAccessToken` in your event-based or spam-reporting add-in. This API is supported in all client versions that support event-based activation, integrated spam reporting, and SSO. On the other hand, `Office.auth.getAccessToken` is only supported in classic Outlook on Windows starting from Version 2111 (Build 14701.20000).

## See also

- [Authenticate a user with a single-sign-on token in an Outlook add-in](../outlook/authenticate-a-user-with-an-sso-token.md)
- [Activate add-ins with events](../develop/event-based-activation.md)
- [Implement an integrated spam-reporting add-in](../outlook/spam-reporting.md)
