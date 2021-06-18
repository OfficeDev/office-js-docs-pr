---
title: Set and get the selected range using the Excel JavaScript API
description: 'Learn how to use the Excel JavaScript API to set and get the selected range using the Excel JavaScript API.'
ms.date: 06/17/2021
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

An add-in can select the edge of the used range, based on the currently selected range. A used range is any cell or series of contiguous cells that have been edited in the worksheet. In the following screenshot, the used range is the table with values in each cell, **C5:F12**. The empty cells outside this table are outside the used range.

![A table with data from C5:F12 in Excel](../images/excel-ranges-used-range.png)

### Select the cell at the edge of the current used range (online-only)

> [!NOTE]
> The `Range.getRangeEdge` method is currently only available in ExcelApiOnline 1.1. To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).

The following code sample shows how use the [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) method to select the cell at the furthest edge of the current used range, in the direction up. This action matches the result of using the Ctrl+Arrow key keyboard shortcut while a range is selected.

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
      activeCell
    );
    rangeEdge.select();

    await context.sync();
}).catch(errorHandlerFunction);
```

> [!NOTE]
> If the selected range is outside a used range, the `Range.getRangeEdge` method will select the cell at the furthest edge of the worksheet in the chosen direction.

#### Before selecting the cell at the edge of the used range

The following screenshot shows a used range and a selected range within the used range. The used range is a table with data at **C5:F12**. Inside this table, the range **D8:E9** is selected. This selection is the *before* state, prior to running the `Range.getRangeEdge` method.

![A table with data from C5:F12 in Excel. The range D8:E9 is selected.](../images/excel-ranges-used-range-d8-e9.png)

#### After selecting the cell at the edge of the used range

The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**. Inside this table, the range **D5** is selected. This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.

![A table with data from C5:F12 in Excel. The range D5 is selected.](../images/excel-ranges-used-range-d5.png)

### Select all cells from current range to furthest edge of used range (online-only)

> [!NOTE]
> The `Range.getExtendedRange` method is currently only available in ExcelApiOnline 1.1. To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).

The following code sample shows how use the [Range.getExtendedRange](javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down. This action matches the result of using the Ctrl+Shift+Arrow key keyboard shortcut while a range is selected.

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
      activeCell
    );
    extendedRange.select();

    await context.sync();
}).catch(errorHandlerFunction);
```

> [!NOTE]
> If the selected range is outside a used range, the `Range.getExtendedRange` method will select all the cells from the current range to the furthest edge of the worksheet in the chosen direction.

#### Before selecting all the cells from the current range to the edge of the used range

The following screenshot shows a used range and a selected range within the used range. The used range is a table with data at **C5:F12**. Inside this table, the range **D8:E9** is selected. This selection is the *before* state, prior to running the `Range.getExtendedRange` method.

![A table with data from C5:F12 in Excel. The range D8:E9 is selected.](../images/excel-ranges-used-range-d8-e9.png)

#### After selecting all the cells from the current range to the edge of the used range

The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**. Inside this table, the range **D8:E12** is selected. This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the up direction.

![A table with data from C5:F12 in Excel. The range D8:E12 is selected.](../images/excel-ranges-used-range-d8-e12.png)

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
