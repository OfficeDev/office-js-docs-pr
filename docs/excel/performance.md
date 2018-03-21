---
title: Excel JavaScript API performance optimization
description: ''
ms.date: 03/13/2017
---

# Performance optimization

Some common tasks can be accomplished via the Excel JavaScript API in more than one way, and there may be significant performance differences between various approaches. This article provides guides and code samples that show how to perform common task in an efficient way using Excel JavaScript API.

## Minimize the number of sync() calls

In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances. To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.

See [Core Concepts - sync()](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts#sync) for code samples that follow this practice.

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
```

## Load necessary properties only

In the Excel JavaScript API, you need to explicitly load the properties and relationships of a proxy object. Although you are able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead. Instead, we strongly suggest that you only load the necessary properties, especially for those objects which have a large number of properties.

For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:
 
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
 
* `properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names. For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).
* `loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://dev.office.com/reference/add-ins/excel/loadoption) for details.

## Suspend calculation temporarily

If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don’t mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.

See [Application Object](https://dev.office.com/reference/add-ins/excel/application) reference documentation for code samples that demonstrate how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.

## Update all cells in a range 

When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array full of a single value. That approach requires Excel to iterate over all of the cells in the range to set each one separately, but Excel has a more efficient way to update all the cells in a range with the same value or property.

So if you’re trying to apply the same value, the same number format or the same formula to a range of cells, it’s much better to use a single value instead of an array of values. This will help you gain a huge performance improvement. For an example code sample that shows this approach in action, see [Update all cells in a range](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts#update-all-cells-in-a-range) in [Core Concepts](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts).

A very common scenario is to apply different number formats on different columns. In this case, you just need to loop over columns and set the number format on each column with a single value. Handle each column as a range as shown in the [Update all cells in a range](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts#update-all-cells-in-a-range) code sample.

## Convert between Range and Table

If you’re trying to import a huge amount of data into a Table object directly (for example, using ```TableRowCollection.add()```),  you might experience a slow performance. You could try to write the data into a range object via ```table.getDataBodyRange()```, and the table will expand automatically.

> NOTE:
> you can conveniently convert between Table objects and Range objects using the [convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange) method of the [Table Object](https://dev.office.com/reference/add-ins/excel/table).

## See also

- [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md)
- [Excel JavaScript API advanced concepts](excel-add-ins-advanced-concepts.md)
- [Excel JavaScript API Open Specification](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Worksheet Functions Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/functions)