---
title: Method element in the manifest file
description: The Method element specifies an individual method from the Office JavaScript API that your Office Add-ins requires in order to activate.
ms.date: 03/19/2019
localization_priority: Normal
---

# Method element

Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.

**Add-in type:** Content, Task pane

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

The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

> [!IMPORTANT]
> Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in. For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).
