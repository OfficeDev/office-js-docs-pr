---
title: Select or get the current Excel range with JavaScript
description: Learn how to select the current Excel range, read the active selection, and extend a selection to the edge of used data in an Office Add-in.
ms.date: 06/03/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Select or get the current Excel range with the JavaScript API

Use these patterns when your Excel add-in needs to move the user's selection, read the currently selected cells, or extend a selection. This article shows the most common selection tasks and links to related range articles when you need to get, format, or update worksheet data.

If you need to get a range by address or by used range instead of by current selection, see [Get Excel worksheet ranges with the JavaScript API](excel-add-ins-ranges-get.md). If you need to read or write the contents of the selected cells, see [Set or get Excel range values, text, and formulas](excel-add-ins-ranges-set-get-values.md). If you need to change how the selected cells look, see [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md).

For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Select a known range

Use `getRange(address)` and `select()` when you know which cells the add-in should highlight. The following example selects **B2:E6** in the active worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("B2:E6");

    range.select();

    await context.sync();
});
```

### Selection after the range is set

:::image type="content" source="../images/excel-ranges-set-selection.png" alt-text="Selected range in Excel.":::

## Get the current selected range

Use `workbook.getSelectedRange()` when the add-in should work with the cells the user already highlighted. In the following example, the add-in gets the selected range, loads its `address` property, and writes the result to the console.

```js
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");

    await context.sync();

    console.log(`The address of the selected range is "${range.address}"`);
});
```

## Select to the edge of a used range

Use [Range.getRangeEdge](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1)) and [Range.getExtendedRange](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1)) when your add-in should match Excel keyboard selection shortcuts. Both methods start from the current selection and move within the used range. To learn more about used ranges, see [Get the used range](excel-add-ins-ranges-get.md#get-the-used-range).

In the following screenshot, the used range is the table that contains values in **C5:F12**. The empty cells outside that table aren't part of the used range.

:::image type="content" source="../images/excel-ranges-used-range.png" alt-text="A table with data from C5:F12 in Excel.":::

### Select the cell at the edge of the used range

Use `Range.getRangeEdge` when you want to move the active selection to the farthest cell in one direction. The following example moves the selection upward to match <kbd>Ctrl</kbd>+<kbd>Up arrow</kbd>.

```js
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const direction = Excel.KeyboardDirection.up;
    const activeCell = context.workbook.getActiveCell();

    const rangeEdge = range.getRangeEdge(direction, activeCell);
    rangeEdge.select();

    await context.sync();
});
```

#### Before selecting the edge cell

The following screenshot shows the used range **C5:F12** with **D8:E9** selected before the `Range.getRangeEdge` method runs.

:::image type="content" source="../images/excel-ranges-used-range-d8-e9.png" alt-text="A table with data from C5:F12 in Excel. The range D8:E9 is selected.":::

#### After selecting the edge cell

The following screenshot shows the same table after the `Range.getRangeEdge` method runs in the up direction. Cell **D5** is selected.

:::image type="content" source="../images/excel-ranges-used-range-d5.png" alt-text="A table with data from C5:F12 in Excel. The range D5 is selected.":::

### Extend the current selection to the edge of the used range

Use `Range.getExtendedRange` when you want to keep the current selection and extend it to the farthest cell in one direction. The following example extends the selection downward to match <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Down arrow</kbd>.

```js
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const direction = Excel.KeyboardDirection.down;
    const activeCell = context.workbook.getActiveCell();

    const extendedRange = range.getExtendedRange(direction, activeCell);
    extendedRange.select();

    await context.sync();
});
```

#### Before extending the selection

The following screenshot shows the used range **C5:F12** with **D8:E9** selected before the `Range.getExtendedRange` method runs.

:::image type="content" source="../images/excel-ranges-used-range-d8-e9.png" alt-text="A table with data from C5:F12 in Excel. The range D8:E9 is selected.":::

#### After extending the selection

The following screenshot shows the same table after the `Range.getExtendedRange` method runs in the down direction. Range **D8:E12** is selected.

:::image type="content" source="../images/excel-ranges-used-range-d8-e12.png" alt-text="A table with data from C5:F12 in Excel. The range D8:E12 is selected.":::

## See also

- [Core Excel object model concepts for Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with Excel cells by using Range objects](excel-add-ins-cells.md)
- [Get Excel worksheet ranges with the JavaScript API](excel-add-ins-ranges-get.md)
- [Set or get Excel range values, text, and formulas](excel-add-ins-ranges-set-get-values.md)
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
