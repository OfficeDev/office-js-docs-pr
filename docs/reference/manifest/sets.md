---
title: Sets element in the manifest file
description: The Sets element specifies the minimum set of Office JavaScript API your Office Add-in requires in order to be activated by Office or to override base manifest settings.
ms.date: 01/22/2022
ms.localizationpriority: medium
---

# Sets element

The meaning of this element depends on where it's used in the manifest.

## In the base manifest

When used in the base manifest (that is, the parent **Requirements** element is a direct child of [OfficeApp](officeapp.md)), the **Sets** element specifies the minimum subset of the Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)) that your Office Add-in needs in order to be activated by Office.

**Add-in type:** Content, Task pane, Mail

## As a grandchild of a VersionOverrides element

Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)) that must be supported by the Office version and platform (such as Windows, Mac, web, and iOS or iPad) in order for the [VersionOverrides](versionoverrides.md) to take effect.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Same as the parent [Requirements](requirements.md) element.

**Associated with these requirement sets**:

- Same as the parent [Requirements](requirements.md) element.

## Syntax

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## Contained in

[Requirements](requirements.md)

## Can contain

[Set](set.md)

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|optional|Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".|

## Remarks

For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Specify which Office versions and platforms can host your add-in](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in).

