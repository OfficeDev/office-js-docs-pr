---
title: EquivalentAddin element in the manifest file
description: ''
ms.date: 04/22/2019
localization_priority: Normal
---

# EquivalentAddin element

specifies backwards compatibility for an equivalent COM add-in, or XLL.

**Add-in type:** Task pane, Custom function

## Syntax

```XML
<EquivalentAddin>  
    <ProgID>{progid}</ProgID>  
    <Type>COM</Type>  
</EquivalentAddin>  
```

## Contained in

[EquivalentAdd-ins](equivalentaddins.md)

## Must contain

[Type](type.md)

## Can contain

[ProgID](progid.md)
[FileName](filename.md)

## Remarks

To specify a COM add-in as the equivalent add-in, provide both the `ProgID` and `Type` elements. To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.

## See also

- [Make your Excel add-in backwards compatible with an existing COM add-in or Excel XLL](/office/dev/add-ins/excel/make-your-excel-add-in-backwards-compatible-with-com-add-in-or-xll)