---
title: Read or write to large ranges using the Excel JavaScript API
description: Learn how to read or write to large ranges with the Excel JavaScript API.
ms.date: 04/02/2021
ms.topic: best-practice
ms.localizationpriority: medium
---

# Read or write to a large range using the Excel JavaScript API

This article describes how to handle reading and writing to large ranges with the Excel JavaScript API.

## Run separate read or write operations for large ranges

If a range contains a large number of cells, values, number formats, or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.

For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).

### Conditional formatting of ranges

Ranges can have formats applied to individual cells based on conditions. For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Read or write to an unbounded range using the Excel JavaScript API](excel-add-ins-ranges-unbounded.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
