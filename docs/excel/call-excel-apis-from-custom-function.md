---
title: Call Excel JavaScript APIs from a custom function
description: Call Excel JavaScript APIs from a custom function to read workbook data safely by using a shared runtime and Excel.RequestContext.
ms.date: 08/18/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Call Excel JavaScript APIs from a custom function

Call Excel JavaScript APIs from a custom function when its calculation needs workbook data that isn't passed in as a parameter. For example, a custom function can read document properties, range values or formats, custom XML parts, or the workbook name.

Keep these calls read-only. A custom function that changes other cells or the Excel environment can cause poor performance, timeouts, or infinite calculation loops.

## Before you begin

Your add-in must use a [shared runtime](../testing/runtimes.md#shared-runtime) before a custom function can call Excel JavaScript APIs. Create an Excel custom functions project with [Microsoft 365 Agents Toolkit](../develop/agents-toolkit-overview.md), or [configure an existing add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## Get a value from the workbook

Create an [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) object to access the workbook. Load the properties that the function needs, call `context.sync()`, and then return the result.

The following custom function uses [Worksheet.getRange](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) to read a cell value. The `address` parameter must be a string. For example, enter `=CONTOSO.GETRANGEVALUE("A1")` in a cell to return the value from cell A1.

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
    // Retrieve the context object.
    const context = new Excel.RequestContext();

    // Use the context object to access the cell at the input address.
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.load("values");
    await context.sync();

    // Return the value of the cell at the input address.
    return range.values[0][0];
}
```

## Limitations of calling Excel JavaScript APIs through a custom function

A custom functions add-in can call Excel JavaScript APIs, but you should be cautious about which APIs it calls. Don't call Excel JavaScript APIs from a custom function that change cells outside of the cell running the custom function. Changing other cells or the Excel environment can result in poor performance, time outs, and infinite loops in the Excel application. This means your custom functions shouldn't do any of the following:

- Insert, delete, or format cells on the spreadsheet.
- Change another cell's value.
- Move, rename, delete, or add sheets to a workbook.
- Add names to a workbook.
- Set properties.
- Change any of the Excel environment options, such as calculation mode or screen views.

Your custom functions add-in can read information from cells outside the cell running the custom function, but it shouldn't perform write operations to other cells. Instead, make changes to other cells or to the Excel environment from the context of a ribbon button or a task pane. In addition, custom function calculations shouldn't run while an Excel recalculation is taking place, as this scenario creates unpredictable results.

## Next steps

- [Excel JavaScript API reference](/javascript/api/excel)

## See also

- [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
