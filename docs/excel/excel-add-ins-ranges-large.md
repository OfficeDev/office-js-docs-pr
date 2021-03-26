---
title: Read or write to large ranges using the Excel JavaScript API
description: 'Code samples that show how to read or write to large ranges using the Excel JavaScript API.'
ms.date: 03/26/2021
localization_priority: Normal
---

# Work with ranges using the Excel JavaScript API

## Read or write to a large range

If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.

For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).

### Conditional formatting of ranges

Ranges can have formats applied to individual cells based on conditions. For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).

## Find a cell using string matching

The `Range` object has a `find` method to search for a specified string within the range. It returns the range of the first cell with matching text. The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console. Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range. If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

When the `find` method is called on a range representing a single cell, the entire worksheet is searched. The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.

## See also

- [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
