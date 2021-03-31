---
title: Work with cells using the Excel JavaScript API.
description: 'Learn the Excel JavaScript API definition of a cell, and learn how to work with cells.'
ms.date: 03/30/2021
localization_priority: Normal
---

# Work with cells using the Excel JavaScript API

The Excel JavaScript API doesn't have a "Cell" object or class. Instead, the Excel JavaScript API defines all Microsoft Excel cells as `Range` objects. An individual Microsoft Excel cell is a `Range` object with one cell.

A single `Range` object can also contain multiple contiguous cells. Contiguous cells are those which form a row, a column, or both. Contiguous cells form an unbroken line. See [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object) to learn about working with cells that are not contiguous.

For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class)](/javascript/api/excel/excel.range).

## Excel JavaScript APIs that apply to cells

CellBorder
CellBorderCollection
CellProperties
CellPropertiesBorderLoadOptions
CellPropertiesFill
CellPropertiesFillLoadOptions
CellPropertiesFont
CellPropertiesFontLoadOptions
CellPropertiesFormat
CellPropertiesFormatLoadOptions
CellPropertiesLoadOptions
CellPropertiesProtection
CellValueConditionalFormat
ConditionalCellValueRule
SettableCellProperties

## Work with discontiguous cells using the RangeAreas object

The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once. These ranges may be contiguous, but do not have to be. `RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

## See also

- [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
