---
title: Find a string using the Excel JavaScript API
description: Learn how to find a string within a range using the Excel JavaScript API.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Find a string within a range using the Excel JavaScript API

This article provides a code sample that finds a string within a range using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Match a string within a range

The `Range` object has a `find` method to search for a specified string within the range. It returns the range of the first cell with matching text.

The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console. Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range. If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let table = sheet.tables.getItem("ExpensesTable");
    let searchRange = table.getRange();
    let foundRange = searchRange.find("Food", {
        completeMatch: true, // Match the whole cell value.
        matchCase: false, // Don't match case.
        searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address");
    await context.sync();

    console.log(foundRange.address);
});
```

When the `find` method is called on a range representing a single cell, the entire worksheet is searched. The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Find special cells within a range using the Excel JavaScript API](excel-add-ins-ranges-special-cells.md)
