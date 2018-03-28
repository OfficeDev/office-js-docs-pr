---
title: Excel JavaScript API performance optimization
description: 'Optimize performance using Excel JavaScript API'
ms.date: 03/28/2018
---

# Performance optimization using the Excel JavaScript API

There are multiple ways that you can perform common tasks with the Excel JavaScript API. You'll find significant performance differences between various approaches. This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.

## Minimize the number of sync() calls

In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online. To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.

See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.

## Minimize the number of proxy objects created

Avoid repeatedly creating the same proxy object. Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.

```javascript
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

In the Excel JavaScript API, you need to explicitly load the properties of a proxy object. Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead. Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.

For example, if you only intend to read the **address** property of a range object, specify only that property when you call the **load()** method:
 
```js
range.load('address');
```
 
You can call **load()** method in any of the following ways:
 
_Syntax:_
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
_Where:_
 
* `properties` is the list of properties to load, specified as comma-delimited strings or as an array of names. For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).
* `loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://dev.office.com/reference/add-ins/excel/loadoption) for details.

Please be aware that some of the “properties” under an object may have the same name as another object. For example, `format` is a property under range object, but `format` itself is an object as well. So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously. To avoid this, your code should only load the “leaf nodes” in an object tree. 

## Suspend calculation temporarily

If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.

See [Application Object](https://dev.office.com/reference/add-ins/excel/application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way. The following code demonstrates how to suspend calculation temporarily:

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

    // Suspending recalc
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with supend recalc
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

## Update all cells in a range 

When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately. Excel has a more efficient way to update all the cells in a range with the same value or property.

If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values. Doing so will significantly improve performance. For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet. In this case, you can simply iterate through the columns and set the number format on each column with a single value. Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.

> [!NOTE]
> If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.  This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.  However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.

## Importing data into tables

When trying to import a huge amount of data directly into a [Table](https://dev.office.com/reference/add-ins/excel/table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance. If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range. If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically. 

Here is an example of this approach:

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first 
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
> You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange) method.

## See also

- [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md)
- [Excel JavaScript API advanced concepts](excel-add-ins-advanced-concepts.md)
- [Excel JavaScript API Open Specification](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Worksheet Functions Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/functions)
