---
ms.date: 03/03/2026
description: Return multiple results from your custom function in an Excel add-in.
title: Return multiple results from your custom function
ms.localizationpriority: medium
---

# Return multiple results from your custom function

Your custom function can return multiple results that fill neighboring cells. This behavior is called "spilling". When a custom function returns an array of results, it's known as a dynamic array formula. This enables your custom functions to work like Excel's built-in dynamic array functions such as `SORT`, `FILTER`, and `UNIQUE`. For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531).

The following image shows how the `SORT` function spills down into neighboring cells. Your custom function can also return multiple results like this.

:::image type="content" source="../images/dynamic-array-spill.png" alt-text="Screen shot of the `SORT` function displaying multiple results down into multiple cells.":::

## Key points

- Return a two-dimensional array to create a custom function that spills results.
- Results spill into neighboring cells automatically.
- If neighboring cells contain data, the formula displays a `#SPILL!` error.
- Arrays spill down by adding rows, right by adding columns, or both for rectangular ranges.
- Dynamic arrays work with streaming functions that update results over time.

## Code samples

To create a custom function that returns dynamic arrays, return a two-dimensional array of values. The array structure determines the spill direction: rows create vertical spills, columns create horizontal spills, and both create rectangular ranges.

### Spill down

The following example returns a dynamic array that spills down. Each inner array represents one row.

```javascript
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}
```

### Spill right

The following example returns a dynamic array that spills right. The single inner array contains multiple values.

```javascript
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}
```

### Spill in both directions

The following example returns a dynamic array that spills both down and right, creating a rectangular range.

```javascript
/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}
```

### Streaming dynamic arrays

Combine dynamic arrays with streaming functions to create results that update over time. The following example returns values that spill down and increment once per second based on the `amount` parameter. To learn more about streaming functions, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).

```javascript
/**
 * Increment the cells with a given amount every second. Creates a dynamic spilled array with multiple results.
 * @customfunction
 * @param {number} amount The amount to add to the cell value on each increment.
 * @param {CustomFunctions.StreamingInvocation<number[][]>} invocation Parameter to send results to Excel or respond to the user canceling the function. A dynamic array.
 */
function increment(amount: number, invocation: CustomFunctions.StreamingInvocation<number[][]>): void {
  let firstResult = 0;
  let secondResult = 1;
  let thirdResult = 2;

  const timer = setInterval(() => {
    firstResult += amount;
    secondResult += amount;
    thirdResult += amount;
    invocation.setResult([[firstResult], [secondResult], [thirdResult]]);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

### Process data with dynamic arrays

Dynamic arrays are useful for processing and transforming input data. The following example takes an array of numbers and returns both the values and their squares.

```javascript
/**
 * Calculate squares of input numbers.
 * @customfunction
 * @param {number[]} numbers Array of numbers to process.
 * @returns {any[][]} A dynamic array showing numbers and their squares.
 */
function calculateSquares(numbers) {
  if (!Array.isArray(numbers)) {
    numbers = [[numbers]];
  }

  // Create header row.
  const result = [['Number', 'Square']];

  // Process each number.
  numbers.forEach(row => {
    const num = Array.isArray(row) ? row[0] : row;
    result.push([num, num * num]);
  });

  return result;
}
```

## See also

- [Dynamic arrays and spilled array behavior](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
- [Custom functions parameter options](custom-functions-parameter-options.md)
- [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function)
- [Handle dynamic arrays and range spilling using the Excel JavaScript API](excel-add-ins-ranges-dynamic-arrays.md)
