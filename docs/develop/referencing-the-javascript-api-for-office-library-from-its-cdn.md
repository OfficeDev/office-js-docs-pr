---
title: Referencing the JavaScript API for Office library in your add-in
description: Learn how to reference the JavaScript API for Office library in your add-in.
ms.date: 02/27/2020
localization_priority: Normal
---

# Referencing the JavaScript API for Office library in your add-in

The [JavaScript API for Office](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office host. The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.

## API versioning and backward compatibility

The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.

> [!IMPORTANT]
> When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.

## Enabling intellisense for a TypeScript project

To enable intellisense for TypeScript add-in project, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder. You must have [Node.js](https://nodejs.org) installed (which includes npm).
> 
> ```command&nbsp;line
> npm install --save-dev @types/office-js
> ```

## See also

- [Understanding the JavaScript API for Office](understanding-the-javascript-api-for-office.md)
- [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office)
