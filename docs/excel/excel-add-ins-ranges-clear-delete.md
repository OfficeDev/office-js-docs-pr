---
title: Clear or delete ranges using the Excel JavaScript API
description: Learn how to clear or delete ranges using the Excel JavaScript API.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Clear or delete ranges using the Excel JavaScript API

This article provides code samples that clear and delete ranges with the Excel JavaScript API. For the complete list of properties and methods supported by the `Range` object, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Key points

- Use `Range.clear` to remove content and formatting from cells without shifting other cells.
- Use `Range.delete` to remove cells and shift surrounding cells to fill the gap.
- `Range.clear` accepts an optional parameter to clear only specific aspects (contents, formats, or hyperlinks).
- `Range.delete` requires specifying the shift direction: up or left.

## Clear a range of cells

The `Range.clear` method removes content and formatting from cells without affecting the worksheet structure. Other cells remain in their current positions. By default, `clear()` removes everything, but you can optionally specify what to clear.

The following code sample clears all contents and formatting of cells in the range **E2:E5**.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("E2:E5");

    range.clear();

    await context.sync();
});
```

### Data before range is cleared

:::image type="content" source="../images/excel-ranges-start.png" alt-text="Data in Excel before range is cleared.":::

### Data after range is cleared

:::image type="content" source="../images/excel-ranges-after-clear.png" alt-text="Data in Excel after range is cleared.":::

### Clear only contents or formatting

Selectively clear only certain aspects of a range using the `applyTo` parameter.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Clear only the contents, keeping formatting intact.
    sheet.getRange("A1:A5").clear(Excel.ClearApplyTo.contents);

    // Clear only the formatting, keeping values intact.
    sheet.getRange("C1:C5").clear(Excel.ClearApplyTo.formats);

    await context.sync();
});
```

## Delete a range of cells

The `Range.delete` method removes cells from the worksheet and shifts surrounding cells to fill the gap. Unlike `clear`, which empties cells, `delete` removes the cells entirely and restructures the worksheet. You must specify which direction to shift the remaining cells: up or left.

The following code sample deletes the cells in the range **B4:E4** and shifts other cells up to fill the space that was vacated by the deleted cells.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    await context.sync();
});
```

### Data before range is deleted

:::image type="content" source="../images/excel-ranges-start.png" alt-text="Data in Excel before range is deleted.":::

### Data after range is deleted

:::image type="content" source="../images/excel-ranges-after-delete.png" alt-text="Data in Excel after range is deleted.":::

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Insert a range of cells using the Excel JavaScript API](excel-add-ins-ranges-insert.md)
- [Cut, copy, and paste ranges using the Excel JavaScript API](excel-add-ins-ranges-cut-copy-paste.md)
