---
title: Referencing the Office JavaScript API library
description: 'Learn how to reference the Office JavaScript API library and type definitions in your add-in.'
ms.date: 02/18/2021
localization_priority: Normal
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

## API versioning and backward compatibility

In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.

> [!NOTE]
> To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## Enabling IntelliSense for a TypeScript project

In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder. You must have [Node.js](https://nodejs.org) installed (which includes npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## Preview APIs

New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is required.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
