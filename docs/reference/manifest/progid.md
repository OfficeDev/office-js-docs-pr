---
title: ProgID element in the manifest file
description: ''
ms.date: 04/22/2019
localization_priority: Normal
---

# ProgID element

Specifies the programmatic identifier of the equivalent COM add-in for the task pane of your web add-in.

**Add-in type:** Task pane

## Syntax

```XML
    <ProgID>{progid}</ProgID>  
```

## Contained in

[EquivalentAdd-in](equivalentaddin.md)

## Remarks

You must specify the programmatic identifier of the COM add-in that contains equivalent UI for your web add-in's task pane UI.

## See also

- [Make your Excel add-in backwards compatible with an existing COM add-in or Excel XLL](/office/dev/add-ins/excel/make-your-excel-add-in-backwards-compatible-with-com-add-in-or-xll.md)