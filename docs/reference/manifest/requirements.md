---
title: Requirements element in the manifest file
description: ''
ms.date: 03/19/2019
localization_priority: Normal
---

# Requirements element

Specifies the minimum set of Office JavaScript API requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<Requirements>
   ...
</Requirements>
```

## Contained in

[OfficeApp](officeapp.md)

## Can contain

|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Methods](methods.md)|x||x|

## Remarks

For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

