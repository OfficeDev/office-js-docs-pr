---
title: PowerPoint JavaScript API requirement sets
description: 'Learn more about the PowerPoint JavaScript API requirement sets.'
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: high
---

# PowerPoint JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.

|  Requirement set  |  Office on Windows<br>(connected to a Microsoft 365 subscription)  |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [PowerPointApi 1.3](powerpoint-api-1-3-requirement-set.md)  | Version 2111 (Build 14701.20060) or later| not yet<br>supported | 16.55 or later | December 2021 |
| [PowerPointApi 1.2](powerpoint-api-1-2-requirement-set.md)  | Version 2011 (Build 13426.20184) or later| not yet<br>supported | 16.43 or later | October 2020 |
| [PowerPointApi 1.1](powerpoint-api-1-1-requirement-set.md) | Version 1810 (Build 11001.20074) or later | 2.17 or later | 16.19 or later | October 2018 |

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## PowerPoint JavaScript API 1.1

PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#PowerPoint_createPresentation_base64File_). For details about the API, see [Create a presentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).

## PowerPoint JavaScript API 1.2

PowerPoint JavaScript API 1.2 adds support for inserting slides from another PowerPoint presentation into the current presentation and for deleting slides. For details about the APIs, see [Insert and delete slides in a PowerPoint presentation](../../powerpoint/insert-slides-into-presentation.md).

## PowerPoint JavaScript API 1.3

PowerPoint JavaScript API 1.3 adds additional support for adding and deleting slides. It also lets add-ins apply custom metadata tags. For details about the APIs, see [Add and delete slides in PowerPoint](../../powerpoint/add-slides.md) and [Use custom tags for presentations, slides, and shapes in PowerPoint](../../powerpoint/tagging-presentations-slides-shapes.md).

## How to use PowerPoint requirement sets at runtime and in the manifest

> [!NOTE]
> This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).

Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.

### Checking for requirement set support at runtime

The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### Defining requirement set support in the manifest

You can use the [Requirements element](/javascript/api/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**. If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that don't support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.

The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.

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
