---
title: Set and get the selected range using the Excel JavaScript API
description: 'Learn how to use the Excel JavaScript API to set and get selected ranges using the Excel JavaScript API.'
ms.date: 06/10/2021
ms.prod: excel
localization_priority: Normal
---

# Set and get the selected range using the Excel JavaScript API

This article provides code samples that set and get the selected range with the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Set the selected range

The following code sample selects the range **B2:E6** in the active worksheet.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### Selected range B2:E6

![Selected range in Excel](../images/excel-ranges-set-selection.png)

## Get the selected range

The following code sample gets the selected range, loads its `address` property, and writes a message to the console.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## Select the edge of a used range

An add-in can select the edges of the used range in the current worksheet, based on the currently selected range. A used range is any cell or series of contiguous cells that have been edited in the worksheet. In the following screenshot, the used range is the table with values in each cell, **C5:F12**. The empty cells outside this table are outside the used range.

![A table with data from C5:F12 in Excel](../images/excel-ranges-used-range.png)

### Select the cell at the edge of the current used range (online-only)

> [!NOTE]
> The `Range.getRangeEdge` method is currently only available in ExcelApiOnline 1.1. To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).

The following code sample shows how to select the cell at the furthest edge of the current used range, in the direction up. This action matches the result of using the Ctrl+Arrow key keyboard shortcut while a range is selected.

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Arrow key keyboard shortcut while a range is selected.
    var rangeEdge = range.getRangeEdge(
      direction,
      activeCell // If the selected range contains more than one cell, the active cell must be defined.
    );
    rangeEdge.select();

    await context.sync();
}).catch(errorHandlerFunction);
```

![A table with data from C5:F12 in Excel. The range D8:E9 is selected.](../images/excel-ranges-used-range-d8-e9.png)

![A table with data from C5:F12 in Excel. The range D5 is selected.](../images/excel-ranges-used-range-d5.png)

### Select all cells from current range to furthest edge of used range (online-only)

> [!NOTE]
> The `Range.getExtendedRange` method is currently only available in ExcelApiOnline 1.1. To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).

The following code sample shows how to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down. This action matches the result of using the Ctrl+Shift+Arrow key keyboard shortcut while a range is selected.

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Arrow key keyboard shortcut while a range is selected.
    var extendedRange = range.getExtendedRange(
      direction,
      activeCell // If the selected range contains more than one cell, the active cell must be defined.
    );
    extendedRange.select();

    await context.sync();
}).catch(errorHandlerFunction);
```

![A table with data from C5:F12 in Excel. The range D8:E9 is selected.](../images/excel-ranges-used-range-d8-e9.png)

![A table with data from C5:F12 in Excel. The range D8:E9 is selected.](../images/excel-ranges-used-range-d8-e12.png)

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
