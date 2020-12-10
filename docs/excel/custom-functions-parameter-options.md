---
ms.date: 12/09/2020
description: 'Learn how to use different parameters within your custom functions, such as Excel ranges, optional parameters, invocation context, and more.'
title: Options for Excel custom functions
localization_priority: Normal
---

# Custom functions parameter options

Custom functions are configurable with many different parameter options.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Optional parameters

When a user invokes a function in Excel, optional parameters appear in brackets. In the following sample, the add function can optionally add a third number. This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.

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
```

---

> [!NOTE]
> When no value is specified for an optional parameter, Excel assigns it the value `null`. This means default-initialized parameters in TypeScript will not work as expected. Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0. Instead use the TypeScript syntax as shown in the previous example.

When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null. In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function. If the `zipCode` parameter is null, the default value is set to `98052`. If the `dayOfWeek` parameter is null, it's set to Wednesday.

#### [JavaScript](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
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
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
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

For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function. 

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## Repeating parameters

A repeating parameter allows a user to enter a series of optional arguments to a function. When the function is called, the values are provided in an array for the parameter. If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`. This matches the convention used for built-in Excel functions.

The following function sums the total of numbers, cell addresses, as well as ranges, if entered.

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
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

### Repeating single value parameter

A repeating single value parameter allows multiple single values to be passed. For example, the user could enter ADD(1,B2,3). The following sample shows how to declare a single value parameter.

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### Single range parameter

A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters. It would appear to the user as ADD(A2:B3) where a single range is passed from Excel. The following sample shows how to declare a single range parameter.

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### Repeating range parameter

A repeating range parameter allows multiple ranges or numbers to be passed. For example, the user could enter ADD(5,B2,C3,8,E5:E8). Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices. For a sample, see the main sample listed for repeating parameters(#repeating-parameters).


### Declaring repeating parameters
In Typescript, indicate that the parameter is multi-dimensional. For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.

In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.

For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.

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
```

## Next steps

Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).

## See also

* [Receive and handle data with custom functions](custom-functions-web-reqs.md)
* [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
* [Manually create JSON metadata for custom functions](custom-functions-json.md)
* [Create custom functions in Excel](custom-functions-overview.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
