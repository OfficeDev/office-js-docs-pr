---
ms.date: 12/07/2022
description: Return multiple results from your custom function in an Excel add-in.
title: Return multiple results from your custom function
ms.localizationpriority: medium
---

# Return multiple results from your custom function

You can return multiple results from your custom function which will be returned to neighboring cells. This behavior is called spilling. When your custom function returns an array of results, it's known as a dynamic array formula. For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531).

The following image shows how the `SORT` function spills down into neighboring cells. Your custom function can also return multiple results like this.

:::image type="content" source="../images/dynamic-array-spill.png" alt-text="Screen shot of the `SORT` function displaying multiple results down into multiple cells.":::

To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values. If the results spill into neighboring cells that already have values, the formula will display a `#SPILL!` error.

## Code samples

The first example shows how to return a dynamic array that spills down.

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

The second example shows how to return a dynamic array that spills right.

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

The third example shows how to return a dynamic array that spills both down and right.

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

The fourth example shows how to return a dynamic spill array from a streaming function. The results spill down, like the first example, and increment once a second based on the `amount` parameter. To learn more about streaming functions, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).

```javascript
/**
 * Increment the cells with a given amount every second. Creates a dynamic spilled array with multiple results
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

## See also

- [Dynamic arrays and spilled array behavior](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Options for Excel custom functions](custom-functions-parameter-options.md)
