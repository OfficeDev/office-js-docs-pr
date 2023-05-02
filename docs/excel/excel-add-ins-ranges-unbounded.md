---
title: Read or write to an unbounded range using the Excel JavaScript API
description: Learn how to use the Excel JavaScript API to read or write to an unbounded range.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Read or write to an unbounded range using the Excel JavaScript API

This article describes how to read and write to an unbounded range with the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

An unbounded range address is a range address that specifies either entire columns or entire rows. For example:

- Range addresses comprised of entire columns.
  - `C:C`
  - `A:F`
- Range addresses comprised of entire rows.
  - `2:2`
  - `1:4`

## Read an unbounded range

When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.

## Write to an unbounded range

You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on an unbounded range because the input request is too large. For example, the following code example is not valid because it attempts to specify `values` for an unbounded range. The API returns an error if you attempt to set cell-level properties for an unbounded range.

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
let range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Read or write to a large range using the Excel JavaScript API](excel-add-ins-ranges-large.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
