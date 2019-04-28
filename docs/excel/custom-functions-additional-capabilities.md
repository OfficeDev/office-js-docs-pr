---
ms.date: 04/25/2019
description: Learn to implement volatile and offline streaming custom functions.
title: Additional capabilities of custom functions (preview)
localization_priority: Normal
---

# Additional capabilities of custom functions

Custom functions can perform mathematical calculations and request information from external sources. More advanced custom functions can start and stop the streaming of data, cancel operations, and handle volatile values within functions.

## Streaming

Custom functions are considered streaming if they request data at set intervals. While it is more common for streaming functions to request web data, they can also regularly perform calculations or other offline actions.

The following example shows a clock function that returns the time each second. This function uses the `invocationContext` parameter, which is always available as the last parameter in any custom function. The function also implements a cancellation handler, which is a best practice when creating a streaming function.

```JavaScript
function clock(invocation) {
  const timer = setInterval() {
    let time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled() {
    clearInterval(timer);
  };
}

CustomFunctions.associate("CLOCK", clock);
```

When you specify metadata for a streaming function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.

```JSON
{
    "id": "CLOCK",
    "name": "CLOCK",
    "description": "Show the current time.",
    "helpUrl": "http://www.contoso.com/help",
    "result": {
        "type": "string",
        "dimensionality": "scalar"
    },
    "parameters": [
    ],
    "options": {
        "cancelable": true,
        "stream": true
    }
}
```

Streaming data from the web requires the same `options` to be declared in the JSON file, but the function's code will instead request data via an XHR or fetch request. For information about streaming web data, see TBD PAGE.

## Offline cancelling

It is a best practice to write a cancellation handler for streaming functions. A cancellation handler can reduce a function's bandwidth consumption, working memory, and CPU load.

Excel automatically cancels a function's execution in the following situations:

- When the user edits or deletes a cell that references the function.

- When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.

- When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.

If you are autogenerating your JSON file, you can declare a cancelable function by using the tag `@cancelable`.

## Volatile values in functions

Volatile functions are functions in which the value changes each time the cell is calculated. The value can change even if none of the function's arguments change. These functions recalculate every time Excel recalculates. For example, imagine a cell that calls the function `NOW`. Every time `NOW` is called, it will automatically return the current date and time.

Excel contains several built-in volatile functions, such as `RAND` and `TODAY`. For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling. For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method
) require the generation of random inputs to determine an optimal solution.

If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
