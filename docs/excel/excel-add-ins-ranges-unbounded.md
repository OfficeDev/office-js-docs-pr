---
title: Read or write to an unbounded range using the Excel JavaScript API
description: Understand what unbounded ranges are, why cell-level operations return null or fail, and how to safely work with entire rows or columns by narrowing scope.
ms.date: 09/22/2025
ms.localizationpriority: medium
---

# Read or write to an unbounded range using the Excel JavaScript API

Use these guidelines to understand how entire-column and entire-row addresses behave, and apply patterns that reduce errors and memory usage. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Key points

- "Unbounded" means entire columns (like `A:F`) or entire rows (such as `2:2`).
- Cell-level properties (like `values`, `text`, `numberFormat`, or `formulas`) come back as `null` for unbounded reads.
- You can't set cell-level properties on an unbounded range. This returns an error.
- Narrow to the used cells first with `getUsedRange()`.
- Prefer explicit bounds (like `A1:F5000`) for faster calculation speeds and lower memory use.

The following are examples of unbounded ranges.

- Range addresses comprised of entire columns.
  - `C:C`
  - `A:F`
- Range addresses comprised of entire rows.
  - `2:2`
  - `1:4`

## Read an unbounded range

When you request an unbounded range (for example, `getRange('C:C')`), the response returns `null` for cell-level properties like `values`, `text`, `numberFormat`, and `formula`. Other properties (`address`, `cellCount`) are still valid.

## Write to an unbounded range

You can't set cell-level properties like `values`, `numberFormat`, or `formula` on an unbounded range because the request is too large. For example, the next code sample fails because it sets `values` for an unbounded range.

```js
// Invalid: Attempting to write cell-level data to unbounded columns.
let range = context.workbook.worksheets.getActiveWorksheet().getRange("A:B");
range.values = [["Due Date"]]; // This throws an error.
```

## Next steps

- Learn strategies for [large bounded ranges](excel-add-ins-ranges-large.md).
- Combine multiple explicit ranges with [multiple ranges](excel-add-ins-multiple-ranges.md).
- Optimize performance with [resource limits guidance](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).
- Identify specific cells using [special cells](excel-add-ins-ranges-special-cells.md).

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Read or write to a large range using the Excel JavaScript API](excel-add-ins-ranges-large.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
