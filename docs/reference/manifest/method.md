---
title: Method element in the manifest file
description: The Method element specifies an individual method from the Office JavaScript API that your Office Add-ins requires in order to be activated by Office or to override base manifest settings.
ms.date: 01/22/2022
ms.localizationpriority: medium
---

# Method element

The meaning of this element depends on where it's used in the manifest.

## In the base manifest

When used in the base manifest (that is, the grandparent **Requirements** element is a direct child of [OfficeApp](officeapp.md)), the **Method** element specifies an individual method from the Office JavaScript API that your Office Add-in needs in order to be activated by Office.

**Add-in type:** Content, Task pane

## As a great-grandchild of a VersionOverrides element

Specifies an individual method from the Office JavaScript API that must be supported by the Office version and platform (such as Windows, Mac, web, and iOS or iPad) in order for the [VersionOverrides](versionoverrides.md) to take effect.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Same as the grandparent [Requirements](requirements.md) element.

**Associated with these requirement sets**:

- Same as the grandparent [Requirements](requirements.md) element.

## Syntax

```XML
<Method Name="string"/>
```

## Contained in

[Methods](methods.md)

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|Name|string|required|Specifies the name of the required method qualified with its parent object. For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.|

## Remarks

The **Methods** and **Method** elements aren't supported by mail add-ins when used in the base manifest. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

> [!IMPORTANT]
> Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in. For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).
