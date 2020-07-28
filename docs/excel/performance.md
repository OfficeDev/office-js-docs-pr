---
title: Excel JavaScript API performance optimization
description: 'Optimize performance using Excel JavaScript API'
ms.date: 07/28/2020
localization_priority: Normal
---

# Performance optimization using the Excel JavaScript API

There are multiple ways that you can perform common tasks with the Excel JavaScript API. You'll find significant performance differences between various approaches. This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.

## Minimize the number of sync() calls

In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web. To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.

## Minimize the number of proxy objects created

Avoid repeatedly creating the same proxy object. Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.

```js
// BAD: repeated calls to .getRange() to create the same proxy object
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: create the range proxy object once and assign to a variable
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

## Load necessary properties only

In the Excel JavaScript API, you need to explicitly load the properties of a proxy object. Although you're able to load all the properties at once with an empty `load()` call, that approach can have significant performance overhead. Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.

For example, if you only intend to read the `address` property of a range object, specify only that property when you call the `load()` method:

```js
range.load('address');
```

You can call `load()` method in any of the following ways:

_Syntax:_

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

_Where:_

* `properties` is the list of properties to load, specified as comma-delimited strings or as an array of names. For more information, see the `load()` methods defined for objects in [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md).
* `loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.

Please be aware that some of the "properties" under an object may have the same name as another object. For example, `format` is a property under range object, but `format` itself is an object as well. So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously. To avoid this, your code should only load the "leaf nodes" in an object tree.

## Suspend Excel processes temporarily

Excel has a number of background tasks reacting to input from both users and your add-in. Some of these Excel processes can be controlled to yield a performance benefit. This is especially helpful when your add-in deals with large data sets.

### Suspend calculation temporarily

If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.

See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way. The following code demonstrates how to suspend calculation temporarily:

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
> You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.

## Untrack unneeded ranges

The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges. These objects persist in memory until `context.sync()` is called. Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.

The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory. Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.

> [!NOTE]
> `Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-). Any proxy object can be untracked by removing it from the tracked objects list in the context. Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.

The following code sample fills a selected range with data, one cell at a time. After the value is added to the cell, the range representing that cell is untracked. Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it. You should notice the code runs faster with the `cell.untrack()` line than without it. You may also notice a quicker response time afterwards, since the cleanup step takes less time.

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // call untrack() to release the range from memory
            cell.untrack();
        }
    }

    await context.sync();
});
```

## See also

* [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
* [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md)
* [Worksheet Functions Object (JavaScript API for Excel)](/javascript/api/excel/excel.functions)
