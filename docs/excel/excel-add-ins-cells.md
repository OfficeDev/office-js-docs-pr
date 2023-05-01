---
title: Work with cells using the Excel JavaScript API
description: Learn the Excel JavaScript API definition of a cell, and learn how to work with cells.
ms.date: 04/16/2021
ms.localizationpriority: medium
---

# Work with cells using the Excel JavaScript API

The Excel JavaScript API doesn't have a "Cell" object or class. Instead, all Excel cells are `Range` objects. An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.

A `Range` object can also contain multiple, contiguous cells. Contiguous cells form an unbroken rectangle (including single rows or columns). To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).

For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).

## Work with discontiguous cells using the RangeAreas object

The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once. These ranges may be contiguous, but they don't have to be. `RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Get a range using the Excel JavaScript API](excel-add-ins-ranges-get.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
