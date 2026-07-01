---
title: Work with Excel cells by using Range objects
description: Learn how Excel cells map to `Range` objects in Office Add-ins, how to target one cell or a block of cells, and when to use `RangeAreas`.
ms.date: 06/03/2026
ms.topic: concept-article
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Work with Excel cells by using `Range` objects

When your Excel add-in needs to read, write, or format a cell, work with a `Range` object. The Excel JavaScript API doesn't include a "Cell" class. Instead, one cell in the Excel UI maps to a `Range` that contains one cell. This article explains that mapping and shows a simple single-cell example.

For the complete list of properties and methods that `Range` supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Work with one cell

In this example, the add-in gets cell **C3**, writes a sales total, makes the value bold, and then reads the address back from Excel.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const cell = sheet.getRange("C3");

    cell.values = [[1250]];
    cell.format.font.bold = true;
    cell.load("address");

    await context.sync();

    console.log(`Updated cell: ${cell.address}`);
});
```

If your add-in needs to read or write the data in a cell or range, see [Set or get Excel range values, text, and formulas](excel-add-ins-ranges-set-get-values.md). If you first need to locate a range, see [Get Excel worksheet ranges with the JavaScript API](excel-add-ins-ranges-get.md).

## Work with multiple contiguous cells

A `Range` can represent more than one cell as long as the cells form one unbroken rectangle. Use a single `Range` when your add-in works with a row, a column, or any rectangular block of adjacent cells. The same patterns for reading and writing to a single cell apply to a block of cells. For example, `Range.values` can read or write to one cell or a block of cells.

## Work with discontiguous cells by using `RangeAreas`

Use the [Excel.RangeAreas](/javascript/api/excel/excel.rangeareas) object when your add-in needs to perform the same operation on multiple ranges at once and those ranges don't touch. For example, `RangeAreas` is the right choice when you need to format **A1**, **C3**, and **E5:E7** together. To learn more, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

## Related articles

- [Core Excel object model concepts](excel-add-ins-core-concepts.md)
- [Get Excel worksheet ranges with the JavaScript API](excel-add-ins-ranges-get.md)
- [Set or get Excel range values, text, and formulas](excel-add-ins-ranges-set-get-values.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
