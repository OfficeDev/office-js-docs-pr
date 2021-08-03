---
title: Excel JavaScript API performance optimization
description: 'Optimize Excel add-in performance using the JavaScript API.'
ms.date: 08/03/2021
localization_priority: Normal
---

# Performance optimization using the Excel JavaScript API

There are multiple ways that you can perform common tasks with the Excel JavaScript API. You'll find significant performance differences between various approaches. This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.

> [!IMPORTANT]
> Many performance issues can be addressed through recommended usage of `load` and `sync` calls. See the "Performance improvements with the application-specific APIs" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) for advice on working with the application-specific APIs in an efficient way.

## Suspend Excel processes temporarily

Excel has a number of background tasks reacting to input from both users and your add-in. Some of these Excel processes can be controlled to yield a performance benefit. This is especially helpful when your add-in deals with large data sets.

### Suspend calculation temporarily

If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.

See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way. The following code demonstrates how to suspend calculation temporarily.

```js
Excel.run(async function(ctx) {
    var app = ctx.workbook.application;
    var sheet = ctx.workbook.worksheets.getItem("sheet1");
    var rangeToSet: Excel.Range;
    var rangeToGet: Excel.Range;
    app.load("calculationMode");
    await ctx.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

Please note that only formula calculations are suspended. Any altered references are still rebuilt. For example, renaming a worksheet still updates any references in formulas to that worksheet.

### Suspend screen updating

Excel displays changes your add-in makes approximately as they happen in the code. For large, iterative data sets, you may not need to see this progress on the screen in real-time. `Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`). Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.

> [!NOTE]
> Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop). Repeated calls will cause the Excel window to flicker.

### Enable and disable events

Performance of an add-in may be improved by disabling events. A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.

## Importing data into tables

When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance. If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range. If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.

Here is an example of this approach:

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    var range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    var table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await ctx.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await ctx.sync();
})
```

> [!NOTE]
> You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#convertToRange__) method.

## Payload size limit best practices

The Excel JavaScript API has size limitations for API calls. Excel on the web has a payload size limit for requests and responses of 5MB, and an API will return a `RichAPI.Error` error if this limit is exceeded. On all platforms, a range is limited to five million cells for get operations. Large ranges typically exceed both of these limitations.

### Payload size

The payload size of a request is a combination of the following three components.

1. The number of API calls.
2. The number of objects, such as a `Range` object.
3. The length of value to be set.

If an API returns the `RequestPayloadSizeLimitExceeded` error, use the best practice strategies documented in this article to optimize your script and avoid the error.

### Strategy 1: Move unchanged values out of loops

In the following code sample, `context.workbook.worksheets.getActiveWorksheet()` can be moved out of the `for` loop, because it doesn't change within that loop.

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    var ranges = [];
    
    for (let i = 0; i < 7500; i++) {
      var rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

The following code sample shows similar logic, but with improved performance. The value `context.workbook.worksheets.getActiveWorksheet()` is retrieved before the `for` loop.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    var ranges = [];
    var worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    for (let i = 0; i < 7500; i++) {
      var rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### Strategy 2: Create fewer range objects

To improve performance and minimize payload size, create fewer range objects.

#### Create fewer range objects: Option A

One way to create fewer range objects is to split each range array into multiple arrays, and then process each new array with multiple loops. *Note: This can reduce the size of each payload request, but using multiple loops will negatively impact performance.*

Reading too many ranges value in one `context.sync()`, the request payload size will exceed the 5MB limitation.

```js
// DO NOT DO THIS. This code sample exceeds the 5MB payload size limit. 
async function run() {
  await Excel.run(async (context) => {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();

    for (let row = 1; row < 10000; row++) {
        var range = sheet.getRangeByIndexes(i, 1, 1, 1);
        range.values = [["1"]];
    }

    await context.sync(); 
  });
}
```

Instead, split the ranges to multiple calls.

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops can negatively impact performance. 
async function run() {
  await Excel.run(async (context) => {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();

    for (let row = 1; row < 5000; row++) {
      var range = worksheet.getRangeByIndexes(i, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
        var range = worksheet.getRangeByIndexes(i, 1, 1, 1);
        range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### Create fewer range objects: Option B

Read all the data in one range and set it in one API call. This will benefit both performance and payload size. Instead of calling `range.values` for each cell in the loop, set the values in an array first, and then call `range.values` in one call.

```js
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    const array = new Array(10000);

    for(var i=0; i<10000; i++) {
      array[i] = [1];
    }

    var range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## See also

* [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
* [Error handling with the Excel JavaScript API](excel-add-ins-error-handling.md)
* [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md)
* [Worksheet Functions Object (JavaScript API for Excel)](/javascript/api/excel/excel.functions)