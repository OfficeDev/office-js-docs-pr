---
title: Calling built-in Excel worksheet functions using the Excel JavaScript API
description: ''
ms.date: 12/19/2019
localization_priority: Normal
---

# Call built-in Excel worksheet functions

This article explains how to call built-in Excel worksheet functions such as `VLOOKUP` and `SUM` using the Excel JavaScript API. It also provides the full list of built-in Excel worksheet functions that can be called using the Excel JavaScript API.

> [!NOTE]
> For information about how to create *custom functions* in Excel using the Excel JavaScript API, see [Create custom functions in Excel](custom-functions-overview.md).

## Calling a worksheet function

The following code snippet shows how to call a worksheet function, where `sampleFunction()` is a placeholder that should be replaced with the name of the function to call and the input parameters that the function requires. The **value** property of the **FunctionResult** object that's returned by a worksheet function contains the result of the specified function. As this example shows, you must `load` the **value** property of the **FunctionResult** object before you can read it. In this example, the result of the function is simply being written to the console.

```js
var functionResult = context.workbook.functions.sampleFunction();
functionResult.load('value');
return context.sync()
    .then(function () {
        console.log('Result of the function: ' + functionResult.value);
    });
```

> [!TIP]
> See the [Excel.Functions](/javascript/api/excel/excel.functions) reference documentation for a completed list of functions that can be called using the Excel JavaScript API.

## Sample data

The following image shows a table in an Excel worksheet that contains sales data for various types of tools over a three month period. Each number in the table represents the number of units sold for a specific tool in a specific month. The examples that follow will show how to apply built-in worksheet functions to this data.

![Screenshot of sales data in Excel for Hammer, Wrench, and Saw in months November, December, and January](../images/worksheet-functions-chaining-results.jpg)

## Example 1: Single function

The following code sample applies the `VLOOKUP` function to the sample data described previously to identify the number of wrenches sold in November.

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
        });
}).catch(errorHandlerFunction);
```

## Example 2: Nested functions

The following code sample applies the `VLOOKUP` function to the sample data described previously to identify the number of wrenches sold in November and the number of wrenches sold in December, and then applies the `SUM` function to calculate the total number of wrenches sold during those two months.

As this example shows, when one or more function calls are nested within another function call, you only need to `load` the final result that you subsequently want to read (in this example, `sumOfTwoLookups`). Any intermediate results (in this example, the result of each `VLOOKUP` function) will be calculated and used to calculate the final result.

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false),
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
        });
}).catch(errorHandlerFunction);
```

## See also

- [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
- [Functions Class (JavaScript API for Excel)](/javascript/api/excel/excel.functions)
- [Workbook Functions Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook#functions)
