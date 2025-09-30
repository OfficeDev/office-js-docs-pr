---
title: Excel JavaScript API performance optimization
description: Optimize Excel add-in performance using the Excel JavaScript API with batching, fewer objects, and reduced payload size.
ms.date: 09/19/2025
ms.topic: best-practice
ms.localizationpriority: medium
---

# Performance optimization using the Excel JavaScript API

Write faster, more scalable Excel add-ins by minimizing processes, batching functions, and reducing payload size. This article shows patterns, anti-patterns, and code samples to help you optimize common operations.

## Quick improvements

Apply these strategies first for the largest immediate impact.

- Batch loads and writes: group property `load` calls, then make a single `context.sync()`.
- Minimize object creation: operate on block ranges instead of many single-cell ranges.
- Write data in arrays, then assign once to a target range.
- Suspend screen updating or calculation only around large changes.
- Avoid per-iteration `Excel.run` or `context.sync()` inside loops.
- Reuse worksheet, table, and range objects instead of re-querying inside loops.
- Keep payloads below size limits by chunking or aggregating before assignment.

> [!IMPORTANT]
> Many performance issues can be addressed through recommended usage of `load` and `sync` calls. See the "Performance improvements with the application-specific APIs" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) for advice on working with the application-specific APIs in an efficient way.

## Suspend Excel processes temporarily

Excel performs background tasks that react to user input and add-in actions. Pausing selected processes can improve performance for large operations.

### Suspend calculation temporarily

If you need to update a large range (such as to assign values and then recalculate dependent formulas) and interim recalculation results aren't needed, suspend calculation temporarily until the next `context.sync()`.

See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way. The following code demonstrates how to suspend calculation temporarily.

```js
await Excel.run(async (context) => {
    let app = context.workbook.application;
    let sheet = context.workbook.worksheets.getItem("sheet1");
    let rangeToSet: Excel.Range;
    let rangeToGet: Excel.Range;
    app.load("calculationMode");
    await context.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await context.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await context.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await context.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
});
```

Only formula calculations are suspended. Any altered references are still rebuilt. For example, renaming a worksheet still updates any references in formulas to that worksheet.

### Suspend screen updating

Excel displays changes as they occur. For large, iterative updates, suppress intermediate screen updates. `Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates until the next `context.sync()` or the end of `Excel.run`. Provide your users with feedback such as status text or a progress bar, because the UI appears idle during suspension.

> [!NOTE]
> Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop). Repeated calls will cause the Excel window to flicker.

### Enable and disable events

You can sometimes improve performance by disabling events. A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.

## Importing data into tables

When you import large datasets directly into a [Table](/javascript/api/excel/excel.table), such as repeatedly calling `TableRowCollection.add()`, performance can degrade. Instead, take the following approach:

1. Write the entire 2D array to a range with `range.values`.
2. Create the table over that populated range (`worksheet.tables.add()`).

For existing tables, set values on `table.getDataBodyRange()` in bulk. The table expands automatically.

Here is an example of this approach:

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    let range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    let table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await context.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await context.sync();
});
```

> [!NOTE]
> You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#excel-excel-table-converttorange-member(1)) method.

## Payload size limit best practices

The Excel JavaScript API has size limitations for API calls. **Excel on the web** limits requests and responses to **5 MB**. The API returns a `RichAPI.Error` error if this limit is exceeded. On all platforms, a range is limited to five million cells for get operations. Large ranges often exceed both limits.

The payload size of a request combines:

- The number of API calls.
- The number of objects, such as `Range` objects.
- The length of the value to set or get.

If you get `RequestPayloadSizeLimitExceeded`, apply the following strategies to reduce size before you split operations.

### Strategy 1: Move unchanged values out of loops

Limit the processes inside loops to improve performance. In the following code sample, `context.workbook.worksheets.getActiveWorksheet()` can be moved out of the `for` loop because it doesn't change within that loop.

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    
    // This sample retrieves the worksheet every time the loop runs, which is bad for performance.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

The following code sample shows similar logic but with an improved strategy. The value `context.workbook.worksheets.getActiveWorksheet()` is retrieved before the loop because it doesn't change. Only values that vary should be retrieved inside the loop.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    // Retrieve the worksheet outside the loop.
    let worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    // Only process the necessary values inside the loop.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### Strategy 2: Create fewer range objects

Create fewer range objects to improve performance and reduce payload size. Two approaches follow.

#### Split each range array into multiple arrays

One way to create fewer range objects is to split each range array into multiple arrays, and then process each new array with a loop and a new `context.sync()` call.

> [!IMPORTANT]
> Only use this strategy after you have confirmed that you exceed the payload size limit. Multiple loops reduce the size of each payload request but also add extra `context.sync()` calls and can hurt performance.

The following code sample attempts to process a large array of ranges in a single loop and then a single `context.sync()` call. Processing too many range values in one `context.sync()` call causes the payload request size to exceed the 5MB limit.

```js
// This code sample does not show a recommended strategy.
// Calling 10,000 rows would likely exceed the 5MB payload size limit in a real-world situation.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // This sample attempts to process too many ranges at once. 
    for (let row = 1; row < 10000; row++) {
      let range = sheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

The following code sample shows logic similar to the preceding code sample, but with a strategy that avoids exceeding the 5MB payload request size limit. In the following code sample, the ranges are processed in two separate loops, and each loop is followed by a `context.sync()` call.

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops and `context.sync()` calls negatively impacts performance.
// Only use this strategy if you've determined that you're exceeding the payload request limit.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();

    // Split the ranges into two loops, rows 1-5000 and then 5001-10000.
    for (let row = 1; row < 5000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    // Sync after each loop. 
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### Set range values in an array

Another way to create fewer range objects is to create an array, use a loop to set all the data in that array, and then pass the array values to a range. This benefits both performance and payload size. Instead of calling `range.values` for each range in a loop, `range.values` is a called once outside the loop.

The following code sample shows how to create an array, set the values of that array in a `for` loop, and then pass the array values to a range outside the loop.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    // Create an array.
    const array = new Array(10000);

    // Set the values of the array inside the loop.
    for (let i = 0; i < 10000; i++) {
      array[i] = [1];
    }

    // Pass the array values to a range outside the loop. 
    let range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## Next steps

- Review [resource limits and performance optimization](../concepts/resource-limits-and-performance-optimization.md) for host-level constraints.
- Explore [working with multiple ranges](excel-add-ins-multiple-ranges.md) to create fewer objects.
- Add telemetry for data such as operation durations and row counts to guide further performance optimization.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Error handling with the application-specific JavaScript APIs](../testing/application-specific-api-error-handling.md)
- [Worksheet Functions Object (JavaScript API for Excel)](/javascript/api/excel/excel.functions)