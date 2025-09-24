---
title: Read or write to large ranges using the Excel JavaScript API
description: Learn strategies to efficiently read or write large Excel ranges with the Excel JavaScript API, without hitting resource limits.
ms.date: 09/22/2025
ms.topic: best-practice
ms.localizationpriority: medium
---

# Read or write to a large range using the Excel JavaScript API

Use these patterns to read or write large ranges, while avoiding resource limit errors.

## Key points

- Don't load or write everything at once. Split big ranges into smaller blocks.
- Load only what you need (for example, just `values` instead of `values,numberFormat,formulas`).
- Choose row blocks or column blocks based on data shape.
- Use `RangeAreas` to work with scattered cells instead of looping each range.
- If you hit a limit error, retry with a smaller block size.
- Apply formatting after the data is in place.

## When to split a large range

| Scenario | Sign you should split the range | Approach |
|----------|----------------------|----------|
| Reading millions of cells | Timeout or resource error | Read in row or column blocks. Start with 5k–20k rows. |
| Writing a large result set | Single `values` write fails | Write in row blocks (with same column count for each block). |
| Sparse updates | Many distant cells | Build combined address string with `getRanges` and `RangeAreas`. |
| Growing table | Don't know final size | Use `getUsedRange()` then step through new rows until empty. |
| Mixed data and formatting | Formatting slows writes | Write values first, format afterward. |

## Defer formatting & calculations

Formatting and calculation-heavy operations, such as conditional formats or formula writes, add time on large areas. Consider:

- First write raw values (plain numbers or text), then add formulas or formats in a second pass.
- Use `setDirty()` only on necessary recalculation scopes.
- Limit conditional formats to used rows instead of entire column references (such as `A2:A5000` instead of `A:A`).

## Handling conditional formatting on large data

Avoid applying conditional formatting to entire columns if only part of the column has data. Find the used range and apply the rule there:

```js
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();
  usedRange.load("address");
  await context.sync();
  const formattedRange = sheet.getRange(usedRange.address.split("!")[1]);
  // Apply formatting rules to formattedRange instead of whole columns.
});
```

## Next steps

- Learn about related [resource limits and performance optimization](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).
- Handle large but sparse selections with [multiple ranges](excel-add-ins-multiple-ranges.md).
- Compare with patterns for [unbounded ranges](excel-add-ins-ranges-unbounded.md).
- Explore special cell targeting in [find special cells](excel-add-ins-ranges-special-cells.md).

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md)
