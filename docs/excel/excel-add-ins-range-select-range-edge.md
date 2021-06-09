---
title: Select the edge of a used range with the Excel JavaScript API
description: 'Learn how to select the edge of a used range with the Excel JavaScript API.'
ms.date: 06/08/2021
ms.prod: excel
localization_priority: Normal
---

# Select the edge of a used range with the Excel JavaScript API

This article provides examples that show different ways to select the edge of a used range within a worksheet using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Select the cell at the edge of the current used range (online-only)

> [!NOTE]
> The `Range.getRangeEdge` method is currently only available in ExcelApiOnline 1.1. To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).

This sample shows how to select the edges of the used range, based on the currently selected range.

The first type of range edge selection selects the cell at the furthest edge of the current used range, in the directions up or left. This action matches the result of using the Ctrl+Arrow key keyboard shortcut while a range is selected.

```js
Excel.run(function (context) {
    // Get the selected range.
    const range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    const direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    const activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Arrow key keyboard shortcut while a range is selected.
    const rangeEdge = range.getRangeEdge(
      direction,
      activeCell // If the selected range contains more than one cell, the active cell must be defined.
    );
    rangeEdge.select();

    await context.sync();
}).catch(errorHandlerFunction);
```

## Select all cells from current range to furthest edge of used range (online-only)

> [!NOTE]
> The `Range.getExtendedRange` method is currently only available in ExcelApiOnline 1.1. To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).

This sample shows how to select the edges of the used range, based on the currently selected range.

The second type of range edge selection selects all the cells from the currently selected range to the furthest edge of the used range, in the directions right or down. This action matches the result of using the Ctrl+Shift+Arrow key keyboard shortcut while a range is selected.

```js
Excel.run(function (context) {
    // Get the selected range.
    const range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    const direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    const activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Arrow key keyboard shortcut while a range is selected.
    const extendedRange = range.getExtendedRange(
      direction,
      activeCell // If the selected range contains more than one cell, the active cell must be defined.
    );
    extendedRange.select();

    await context.sync();
}).catch(errorHandlerFunction);
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get ranges using the Excel JavaScript API](excel-add-ins-ranges-set-get.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
