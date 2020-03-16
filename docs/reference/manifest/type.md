---
title: Type element in the manifest file
description: The Type element specifies if the equivalent add-in is a COM add-in or an XLL.
ms.date: 03/16/2020
localization_priority: Normal
---

# Type element

Specifies if the equivalent add-in is a COM add-in or an XLL.

**Add-in type:** Task pane, Custom function

## Syntax

```XML
    <Type> [COM | XLL] </Type>  
```

## Contained in

[EquivalentAdd-in](equivalentaddin.md)

## Add-in type values

You must specify one of the following values for the `Type` element.

- COM: Specifies the equivalent add-in is a COM add-in.
- XLL: Specifies the equivalent add-in is an Excel XLL.

## See also

- [Make your custom functions compatible with XLL user-defined functions](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Make your Excel add-in compatible with an existing COM add-in](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)