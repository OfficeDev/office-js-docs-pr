---
title: Methods element in the manifest file
description: The Methods element specifies the list of Office JavaScript API methods that your Office Add-in requires in order to be activated by Office or to override base manifest settings.
ms.date: 01/22/2022
ms.localizationpriority: medium
---

# Methods element

The meaning of this element depends on where it's used in the manifest.

## In the base manifest

When used in the base manifest (that is, the parent **Requirements** element is a direct child of [OfficeApp](officeapp.md)), the **Methods** element specifies the list of Office JavaScript API methods that your Office Add-in needs in order to be activated by Office.

**Add-in type:** Content, Task pane

## As a grandchild of a VersionOverrides element

Specifies the minimum set of Office JavaScript API methods that must be supported by the Office version and platform (such as Windows, Mac, web, and iOS or iPad) in order for the [VersionOverrides](versionoverrides.md) to take effect.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Same as the parent [Requirements](requirements.md) element.

**Associated with these requirement sets**:

- Same as the parent [Requirements](requirements.md) element.

## Syntax

```XML
<Methods>
   ...
</Methods>
```

## Contained in

[Requirements](requirements.md)

## Can contain

[Method](method.md)

## Remarks

The **Methods** and **Method** elements aren't supported in mail add-ins when used in the base manifest. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).
