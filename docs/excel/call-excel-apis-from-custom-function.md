---
title: Call Excel JavaScript APIs from a custom function
description: Learn which Excel JavaScript APIs you can call from your custom function.
ms.date: 03/20/2023
ms.localizationpriority: medium
---

# Call Excel JavaScript APIs from a custom function

Call Excel JavaScript APIs from your custom functions to get range data and obtain more context for your calculations. Calling Excel JavaScript APIs through a custom function can be helpful when:

- A custom function needs to get information from Excel before calculation. This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.
- A custom function will set the cell's number format for the return values after calculation.

> [!IMPORTANT]
> To call Excel JavaScript APIs from your custom function, you'll need to use a [shared runtime](../testing/runtimes.md#shared-runtime). Use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to install an **Excel Custom Functions using a Shared Runtime** project or see [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.

## Code sample

To call Excel JavaScript APIs from a custom function, you first need a context. Use the [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) object to get a context. Then use the context to call the APIs you need in the workbook.

The following code sample shows how to use `Excel.RequestContext` to get a value from a cell in the workbook. In this sample, the `address` parameter is passed into the Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) method and must be entered as a string. For example, the custom function entered into the Excel UI must follow the pattern `=CONTOSO.GETRANGEVALUE("A1")`, where `"A1"` is the address of the cell from which to retrieve the value.

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

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)

## See also

- [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
