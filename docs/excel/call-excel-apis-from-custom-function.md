---
title: Call Microsoft Excel APIs from a custom function
description: 'Learn which Microsoft Excel APIs you can call from your custom function.'
ms.date: 02/06/2020
localization_priority: Normal
---

# Call Microsoft Excel APIs from a custom function

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

Call Office.js Excel APIs from your custom functions to get range data and obtain more context for your calculations.

Calling Office.js APIs through a custom function can be helpful when:

- A custom function needs to get information from Excel before calculation. This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.
- A custom function will set the cell's number format for the return values after calculation.

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## Code sample

To call into the Office.js APIs you first need a context. Use the `Excel.RequestContext` object to get a context. Then use the context to call the APIs you need in the workbook.

The following code sample shows how to get a range of values from the workbook.

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

## Limitations of calling Office.js through a custom function

Don't call Office.js APIs from a custom function that change the environment of Excel. This means your custom functions should not do any of the following:

- Insert, delete, or format cells on the spreadsheet.
- Change another cell's value.
- Move, rename, delete, or add sheets to a workbook.
- Change any of the environment options, such as calculation mode or screen views.
- Add names to a workbook.
- Set properties or execute most methods.

Changing Excel can result in poor performance, time outs, and infinite loops. Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.

Instead, make changes to Excel from the context of a ribbon button, or task pane.

## Next steps

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)

## See also

- [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)