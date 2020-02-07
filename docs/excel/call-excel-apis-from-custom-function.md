---
title: Call Excel APIs from a custom function
description: Learn which Excel APIs you can call from your custom function. 
ms.date: 02/06/2020
localization_priority: Normal
---

# Call Excel APIs from a custom function

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

Call Office.js Excel APIs from your custom functions to get range data and obtain more context for your calculations.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Calling Office.js APIs through a custom function can be helpful when:

- A custom function needs to get information from Excel before calculation. This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.
- A custom function will set the cell's number format for the return values after calculation.

## Code sample

To call into the Office.js APIs you first need a context. Use the `Excel.RequestContext` object to get a context. Then use the context to call the APIs you need in the workbook.

The following code sample shows how to get a range of values from the workbook.

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

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

Don't call Office.js APIs from a custom funciton that write or set information. These APIs include actions such as adding sheets to a workbook, changing cell values, or formatting cells on a spreadsheet. Writing or setting data can result in poor performance, time outs, and infinite loops. Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.

Instead, call Office.js APIs first before a custom function is run. Another option is to use parameters within your custom function to pass information which is being set or written.

## Next steps

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)

## See also

- [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)