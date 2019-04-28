---
ms.date: 04/28/2019
description: Learn how to use different parameters within your custom functions, such as Excel ranges, optional parameters, invocation context, and more.   
title: Options for Excel custom functions (preview)
localization_priority: Normal
---

# Custom functions parameter options

Custom functions are configurable with many different options for parameters

## Optional parameters

Whereas regular parameters are required, optional parameters are not. When a user invokes a function in Excel, optional parameters appear in brackets. In the following sample, the add function can optionally add a third number. This function would appear as `=CONTOSO.ADD(first, second, [third])` in Excel.

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number
 * @param {number} second Second number
 * @param {number} [third] Third number
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, [third]) {
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined. In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function. If the `zipCode` parameter is undefined, the default value is set to `98052`. If the `dayOfWeek` parameter is undefined, it is set to Wednesday.

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code
 * @param {string} dayOfWeek Day of the week
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## Range parameters

Your custom function may accept a range of data as an input parameter. A function can also return a range of data. In JavaScript, a range of data is represented as a two-dimensional array.

For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`. Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## Invocation context parameter

Every custom function is automatically passed an `invocation` argument as the last argument, which can be used when you wish to find the address of a cell or handle what happens when [canceling a function](custom-functions-web-reqs.md#stream-and-cancel-functions). Even if you declare no parameters, your custom function has this parameter. This argument doesn't appear for a user in Excel. If you want to use `invocation` in your custom function, declare it as the last parameter.

In the following code sample, the `invocation` context is explicitly stated for your reference.

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).

### Addressing cell's context parameter

In some cases you need to get the address of the cell that invoked your custom function. This is useful in the following types of scenarios:

- Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.
- Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.
- Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.

To request an addressing cell's context in a function, you need to use a helper function to find the cell's address, such as the one in the following example. The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.

```js
/**
 * Helper function to get the address of a cell
 * @customfunction
 * @param invocation Uses the invocation parameter present in each cell
 * @requiresAddress 
 * @returns {string} Returns address of cell
 */

function getAddress(invocation) {
    return invocation.address;
}
```

By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`. For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
