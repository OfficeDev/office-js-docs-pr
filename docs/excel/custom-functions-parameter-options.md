---
ms.date: 04/23/2019
description: Learn how to use different parameters within your custom functions, such as Excel ranges, optional parameters, invocation context, and more.   
title: Options for Excel custom functions (preview)
localization_priority: Normal
---
# Custom functions parameter options

Custom functions are configurable with many different options for parameters.

## Optional parameters

Whereas regular parameters are required, optional parameters are not. When a user invokes a function in Excel, optional parameters appear in brackets. 

For example, a function `=MAKEPIZZA` with one required parameter called `cheese` and one optional parameter called `anchovies` would appear as `=MAKEPIZZA(cheese, [anchovies])` in Excel.

To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function. The following example shows what this might look like for the function `=ADD(first, second, [third])`. Notice that the optional `[third]` parameter follows the two required parameters. Required parameters will appear first in Excel’s Formula UI.

```json
{
    "id": "ADD",
    "name": "ADD",
    "description": "Adds two or more numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
        },
    "parameters": [
        {
            "name": "first",
            "description": "first number to add",
            "type": "number",
            "dimensionality": "scalar"
        },
        {
            "name": "second",
            "description": "second number to add",
            "type": "number",
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ]
}
```

When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined. In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function. If the `zipCode` parameter is undefined, the default value is set to 98052. If the `dayOfWeek` parameter is undefined, it is set to Wednesday.

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## Range parameters

Your custom function may accept a range of data as an input parameter, or it may return a range of data. In JavaScript, a range of data is represented as a two-dimensional array.

For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`. Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.

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

Every custom function is automatically passed an `invocationContext` argument as the last argument. Even if you declare no parameters, your custom function will have this parameter. This argument doesn't appear for a user in Excel. If you want to use `invocationContext` in your custom function, declare it as the last parameter.

In the following code sample, you'll see the invocationContext explicitly stated for your reference.

```JavaScript
subtract(one, two, invocationContext) {
  return one - two;
}
```

The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).

## Addressing cell's context parameter

In some cases you'll need to get the address of the cell that invoked your custom function. This may be useful in the following types of scenarios:

- Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Then, use [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.
- Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.
- Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.

The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file. The following sample gives an example of this:

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

In the script file (**./src/functions/functions.js** or **./src/functions/functions.ts**), you'll also need to add a `getAddress` function to find a cell's address. This function may take parameters, as shown in the following sample as `parameter1`. The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`. For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
