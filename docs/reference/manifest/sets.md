---
title: Sets element in the manifest file
description: The Sets element specifies the minimum set of Office JavaScript API your Office Add-in requires in order to activate.
ms.date: 03/19/2019
localization_priority: Normal
---

# Sets element

Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.

**Add-in type:** Content, Task pane, Mail

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

For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

