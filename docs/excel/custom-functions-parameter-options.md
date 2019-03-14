---
ms.date: 03/14/2019
description: Learn how to use different parameters within your custom functions, such as Excel ranges, optional parameters, invocation context, and more.   
title: Options for Excel custom functions (preview)
localization_priority: Normal
---
# Custom functions parameter options

Custom functions are configurable with many different options. This article will cover various TBD

## Optional parameters

TBD

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

TBD

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

In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address. This function may take parameters, as shown in the following sample as `parameter1`. The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.

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
