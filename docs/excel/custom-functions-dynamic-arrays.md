---
ms.date: 12/16/2019
description: Return a dynamic array from your custom function in an Office Excel add-in.
title: Return a dynamic array from your custom function
localization_priority: Normal
---

# Return a dynamic array from your custom function

You can return multiple results from your custom function to make it a dynamic array formula. Multiple results will spill into neighboring cells. You can return results that only spill down, to the right, or both down and to the right.

To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values. The following example shows how to return a dynamic array that spills down.

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

The following example shows how to return a dynamic array that spills right. 

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

The following example shows how to return a dynamic array that spills both down and right.

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

## See also

- [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Options for Excel custom functions](custom-functions-parameter-options.md)