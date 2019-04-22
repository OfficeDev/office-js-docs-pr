---
title: Type element in the manifest file
description: ''
ms.date: 04/22/2019
localization_priority: Normal
---

# Type element

specifies if the equivalent add-in is a COM addin or an XLL.

**Add-in type:** Task pane, Custom function

## Syntax

```XML
    <Type>{Add-in type}</Type>  
```

## Contained in

[EquivalentAdd-in](equivalentaddin.md)

## Add-in type values

You must specify one of the following values for the `Type` element.

- COM: Specifies the equivalent add-in is a COM add-in.
- XLL: Specifies the equivalent add-in is an Excel XLL.

## See also

- [Make your Excel add-in backwards compatible with an existing COM add-in or Excel XLL](/office/dev/add-ins/excel/make-your-excel-add-in-backwards-compatible-with-com-add-in-or-xll.md)