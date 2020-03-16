---
title: Requirements element in the manifest file
description: 'The Requirements element specifies the minimum requirement set and methods your Office Add-in needs in order to activate.'
ms.date: 03/19/2019
localization_priority: Normal
---

# Requirements element

Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.

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

For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).
