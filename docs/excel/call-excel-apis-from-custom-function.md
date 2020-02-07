---
title: Call Excel APIs from a custom function
description: Learn which Excel APIs you can call from your custom function. 
ms.date: 02/06/2020
localization_priority: Normal
---
# Call Excel APIs from a custom function

Custom functions are able to call most Office.js Excel APIs to get range data and obtain more context for your calculations.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Calling these APIs through a custom function can be helpful if:

- A custom function needs to get information from Excel before calculation. This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.

- A custom function will set the cell's number format for the return values after calculation.

## Code sample

The following code sample shows you how to do TODO. This sample will only work if you have made the changes to your manifest and your task pane's HTML file as shown in the [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial).

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## Restrictions to calling Office.js through a custom function

Calling Office.js through a custom function isn't recommended if:

- The API you intend to call **writes** or **sets** information. These APIs include actions such as adding sheets to a workbook, changing cell values, or formatting cells on a spreadsheet.

This can result in poor performance, time outs, and infinite loops. Custom functions' calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.

Instead, call Office.js APIs first before a custom function is run. Another option is to use parameters within your custom function to pass information which is being set or written.

## Next steps

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)

## See also

- [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial)