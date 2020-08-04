---
title: PowerPoint JavaScript API requirement sets
description: 'Learn more about the PowerPoint JavaScript API requirement sets.'
ms.date: 07/10/2020
ms.prod: powerpoint
localization_priority: Priority
---

# PowerPoint JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.

|  Requirement set  |  Office on Windows<br>(connected to a Microsoft 365 subscription)  |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1.1 | Version 1810 (Build 11001.20074) or later | 2.17 or later | 16.19 or later | October 2018 |

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## PowerPoint JavaScript API 1.1

PowerPoint JavaScript API 1.1 contains a single API to create a new presentation. For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).

## Runtime requirement support check

At runtime, add-ins can check if a particular application supports an API requirement set by doing the following.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## Manifest-based requirement support check

Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.

The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## Office Common API requirement sets

Most of the PowerPoint Add-in functionality comes from the Common API set. For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [PowerPoint JavaScript API reference documentation](/javascript/api/powerpoint)
- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
