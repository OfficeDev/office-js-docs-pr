---
title: Set element in the manifest file
description: The Set element specifies an Office JavaScript API requirement set your Office Add-in requires in order to be activated by Office or to override base manifest settings.
ms.date: 01/22/2022
ms.localizationpriority: medium
---

# Set element

The meaning of this element depends on where it's used in the manifest.

## In the base manifest

When used in the base manifest (that is, the grandparent **Requirements** element is a direct child of [OfficeApp](officeapp.md)), the **Set** element specifies a [requirement set](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) from the Office JavaScript API that your Office Add-in needs in order to be activated by Office.

**Add-in type:** Content, Task pane, Mail

## As a great-grandchild of a VersionOverrides element

Specifies a [requirement set](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) from the Office JavaScript API that must be supported by the Office version and platform (such as Windows, Mac, web, and iOS or iPad) in order for the [VersionOverrides](versionoverrides.md) to take effect.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Same as the grandparent [Requirements](requirements.md) element.

**Associated with these requirement sets**:

- Same as the grandparent [Requirements](requirements.md) element.

## Syntax

```XML
<Set Name="string" MinVersion="n .n">
```

## Contained in

[Sets](sets.md)

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|Name|string|required|The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).|
|MinVersion|string|optional|Specifies the minimum version of the API set required by your add-in. Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.|

## Remarks

For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Specify which Office versions and platforms can host your add-in](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in).

