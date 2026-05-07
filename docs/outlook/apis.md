---
title: Outlook add-in APIs
description: Learn how to load the Office.js library, specify Mailbox requirement sets, declare permissions, and work with the Mailbox object in your Outlook add-in.
ms.date: 05/06/2026
ms.topic: overview
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Outlook add-in APIs

Before your Outlook add-in can read message data, update items, or call external services, it needs a reference to the Office.js library, a declared requirement set, and the right permissions. This article explains how to configure each of these building blocks and how to access the Outlook API through the [Mailbox](#mailbox-object) object.

## Office.js library

To interact with the [Outlook add-in API](/javascript/api/outlook), you need to use the JavaScript APIs in Office.js. The content delivery network (CDN) for the library is `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. Add-ins submitted to Microsoft Marketplace must reference Office.js by this CDN; they can't use a local reference.

Reference the CDN in a `<script>` tag in the `<head>` tag of the web page (.html, .aspx, or .php file) that implements the UI of your add-in.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

As we add new APIs, the URL to Office.js will stay the same. We'll change the version in the URL only if we break an existing API behavior.

> [!IMPORTANT]
> When developing an add-in for any Office client application, reference the Office JavaScript API from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements.

## Requirement sets

All Outlook APIs belong to the [Mailbox requirement set](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). The `Mailbox` requirement set has versions, and each new set of APIs that are released belongs to a higher version of the set. Not all Outlook clients will support the newest set of APIs when they're released, but if an Outlook client declares support for a requirement set, it will support all the APIs in that requirement set.

To control which Outlook clients the add-in appears in, specify a minimum requirement set version in the manifest. For example, if you specify requirement set version 1.3, the add-in won't show up in any Outlook client that doesn't support a minimum version of 1.3.

Specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use v1.3 APIs. The requirement set only controls which Outlook clients the add-in appears in.

To check the availability of any APIs from a requirement set greater than the one specified in the manifest, you can use standard JavaScript. For example, the following code checks for `getItemIdAsync`, which requires Mailbox requirement set 1.8, before calling it.

```js
// Check whether a Mailbox 1.8 API is available before calling it.
if (Office.context.mailbox.item.getItemIdAsync) {
    Office.context.mailbox.item.getItemIdAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Item ID:", result.value);
        }
    });
}
```

> [!NOTE]
> These checks are not needed for any APIs that are in the requirement set version specified in the manifest.

Specify the minimum requirement set that supports the critical set of APIs for your scenario, without which features of your add-in won't work. You specify the requirement set in the manifest. The markup varies depending on the manifest that you are using.

- **Add-in only manifest**:  Use the `<Requirements>` element. Note that the `<Methods>` child element of `<Requirements>` isn't supported in Outlook add-ins, so you can't declare support for specific methods.
- **Unified manifest for Microsoft 365**: Use the `"extensions.capabilities"` property.

For more information, see [Office Add-in manifests](../develop/add-in-manifests.md), and [Understanding Outlook API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

## Permissions

Your add-in requires the appropriate permissions to use the APIs that it needs. In general, you should specify the minimum permission needed for your add-in.

There are four levels of permissions: **restricted**, **read item**, **read/write item**, and **read/write mailbox**. For more details, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).

## Mailbox object

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## See also

- [Office Add-in manifests](../develop/add-in-manifests.md)
- [Understanding Outlook API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md)
- [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md)
