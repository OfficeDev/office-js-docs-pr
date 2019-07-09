---
ms.date: 07/09/2019
description: Learn how to use different parameters within your custom functions, such as Excel ranges, optional parameters, invocation context, and more.   
title: Options for Excel custom functions
localization_priority: Normal
---

# Custom functions parameter options

Custom functions are configurable with many different options for parameters.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Optional parameters

Whereas regular parameters are required, optional parameters are not. When a user invokes a function in Excel, optional parameters appear in brackets. In the following sample, the add function can optionally add a third number. This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.

#### [JavaScript](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
CustomFunctions.associate("ADD", add);
```

#### [TypeScript](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
CustomFunctions.associate("ADD", add);
```

---

> [!NOTE]
> When no value is specified for an optional parameter, Excel assigns it the value `null`. This means default-initialized parameters in TypeScript will not work as expected. Therefore, don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0. Instead use the TypeScript syntax as shown in the previous example.

When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are null. In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function. If the `zipCode` parameter is null, the default value is set to `98052`. If the `dayOfWeek` parameter is null, it is set to Wednesday.

#### [JavaScript](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### [TypeScript](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string
{
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## Range parameters

Your custom function may accept a range of cell data as an input parameter. A function can also return a range of data. Excel will pass a range of cell data as a two-dimensional array.

For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`. Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
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
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## Repeating parameters 
A repeating parameter allows a user to enter a series of optional of arguments to a function. When the function is called, the values are provided in an array for the parameter. If the parameter name ends with a number, each argument will increment the number, such as `ADD(number1, [number2], [number3],…)`. This matches the convention used for built-in Excel functions.

The following function sums the total of numbers, cell addresses, as well as ranges, if entered.  

```TS
/**  
* The sum of all of the numbers.  
* @customfunction  
* @param operands A number (such as 1 or 3.1415)  
*/  

function ADD(operands: number[][][]): number {  
  let total: number = 0;  

  operands.forEach(range => {  
    range.forEach(row => {  
      row.forEach(num => {  
        total += num;  
      });  
    });  
  });  

  return total;  
}
```

This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### Declaring repeating parameters
In Typescript, indicate that the parameter is multi-dimensional. For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.  

In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.

For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality”: matrix`.  

>[!NOTE]
>Functions containing repeating parameters automatically contain an invocation parameter as the last parameter. For more information on invocation parameters, see the following section.

## Invocation parameter

Every custom function is automatically passed an `invocation` argument as the last argument. This argument can be used to retrieve additional context, such as the address of the calling cell. Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function). Even if you declare no parameters, your custom function has this parameter. This argument doesn't appear for a user in Excel. If you want to use `invocation` in your custom function, declare it as the last parameter.

In the following code sample, the `invocation` context is explicitly stated for your reference.

```js
/**
 * Add two numbers.
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).

### Addressing cell's context parameter

In some cases you need to get the address of the cell that invoked your custom function. This is useful in the following scenarios:

- Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.
- Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.
- Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.

To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example. The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
CustomFunctions.associate("GETADDRESS", getAddress);
```

By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`. For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.

## Next steps
Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).

## See also

* [Receive and handle data with custom functions](custom-functions-web-reqs.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions metadata](custom-functions-json.md)
* [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
* [Create custom functions in Excel](custom-functions-overview.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
