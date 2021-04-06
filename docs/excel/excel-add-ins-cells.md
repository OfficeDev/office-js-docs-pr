---
title: Work with cells using the Excel JavaScript API.
description: 'Learn the Excel JavaScript API definition of a cell, and learn how to work with cells.'
ms.date: 04/02/2021
localization_priority: Normal
---

# Work with cells using the Excel JavaScript API

The Excel JavaScript API doesn't have a "Cell" object or class. Instead, all Excel cells are `Range` objects. An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.

A `Range` object can also contain multiple, contiguous cells. Contiguous cells form an unbroken rectangle (including single rows or columns). See [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object) to learn about working with cells that are not contiguous.

For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).

## Excel JavaScript APIs that mention cells

Even though the Excel JavaScript API doesn't have a "Cell" object or class, a number of API names mention cells. These APIs tend to control cell properties like color, text formatting, and font.

The following is a list of Excel JavaScript APIs that refer to cells.

- [CellBorder](/javascript/api/excel/excel.cellborder)
- [CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)
- [CellProperties](/javascript/api/excel/excel.cellproperties)
- [CellPropertiesBorderLoadOptions](/javascript/api/excel/excel.cellpropertiesborderloadoptions)
- [CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)
- [CellPropertiesFillLoadOptions](/javascript/api/excel/excel.cellpropertiesfillloadoptions)
- [CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)
- [CellPropertiesFontLoadOptions](/javascript/api/excel/excel.cellpropertiesfontloadoptions)
- [CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)
- [CellPropertiesFormatLoadOptions](/javascript/api/excel/excel.cellpropertiesformatloadoptions)
- [CellPropertiesLoadOptions](/javascript/api/excel/excel.cellpropertiesloadoptions)
- [CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)
- [CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)
- [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)
- [SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)

## Work with discontiguous cells using the RangeAreas object

The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once. These ranges may be contiguous, but they don't have to be. `RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Get a range using the Excel JavaScript API](excel-add-ins-ranges-get.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
