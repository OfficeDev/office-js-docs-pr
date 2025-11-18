---
title: Referencing the Office JavaScript API library
description: Learn how to reference the Office JavaScript API library and type definitions in your add-in.
ms.date: 01/14/2025
ms.localizationpriority: medium
---

# Referencing the Office JavaScript API library

The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office application. The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.

> [!IMPORTANT]
> You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.

## Office.js-specific web API behavior

Office.js replaces the default [Window.history](https://developer.mozilla.org/docs/Web/API/History) methods of `replaceState` and `pushState` with `null`. If your add-in relies on these methods, replace the Office.js library reference with the following workaround.

```HTML
<script type="text/javascript">
    // Cache the history method values.
    window._historyCache = {
        replaceState: window.history.replaceState,
        pushState: window.history.pushState
    };
</script>

<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

<script type="text/javascript">
    // Restore the history method values after loading Office.js
    window.history.replaceState = window._historyCache.replaceState;
    window.history.pushState = window._historyCache.pushState;
</script>

```

Thank you to [@stepper and the Stack Overflow community](https://stackoverflow.com/questions/42642863/office-js-nullifies-browser-history-functions-breaking-history-usage) for suggesting and verifying this workaround.

## API versioning and backward compatibility

In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.

If you plan to publish your Office Add-in from Microsoft Marketplace, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.

> [!NOTE]
> To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## Enabling IntelliSense for a TypeScript project

In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder. You must have [Node.js](https://nodejs.org) installed (which includes npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## Preview APIs

New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is acquired.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## CDN references for other Microsoft 365 environments

[!INCLUDE [Information about the China-specific CDN](../includes/21Vianet-CDN.md)]

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
- [Guidance for deploying Office Add-ins on government clouds](../publish/government-cloud-guidance.md)
- [Microsoft software license terms for the Microsoft Office JavaScript (Office.js) API library](https://github.com/OfficeDev/office-js/blob/release/LICENSE.md)
