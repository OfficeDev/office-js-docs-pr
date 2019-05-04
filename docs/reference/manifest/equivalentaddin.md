---
title: EquivalentAddin element in the manifest file
description: ''
ms.date: 04/22/2019
localization_priority: Normal
---

# EquivalentAddin element

Specifies backwards compatibility for an equivalent COM add-in or XLL.

**Add-in type:** Task pane, Custom function

## Syntax

```XML
<EquivalentAddin>
   ...
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

- [Make your custom functions compatible with XLL user-defined functions](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Make your Excel add-in compatible with an existing COM add-in](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)