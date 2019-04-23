---
ms.date: 04/23/2019
description: Learn to implement volatile and offline streaming custom functions.
title: Offline options for Excel custom functions (preview)
localization_priority: Normal
---
# Custom functions additional capabilities

Custom functions can perform mathematical calculations and request information from external sources, such as the web. However, they also have some additional capabilities including streaming, cancelling, and handling volatile functions.

## Offline streaming

Custom functions are considered streaming if they request data at set intervals. While it is more common for streaming functions to request web data, they can also perform calculations or other offline actions at set intervals.

The following example shows a clock function that returns the time each second. Note that this function uses the invocationContext parameter, which is automatically available as a parameter in your custom function. The function also implements a cancellation handler, which is a best practice when creating a streaming function.

```JavaScript
function clock(invocation) {
  const timer = setInterval() {
    const time = currentTime();
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

Excel automatically cancels the execution of a function in the following situations:

- When the user edits or deletes a cell that references the function.

- When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.

- When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.

To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function. The code samples in the previous section of this article provide an example of these techniques.

## Volatile values in functions

[Volatile functions](https://docs.microsoft.com/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed. These functions recalculate every time Excel recalculates. For example, imagine a cell that calls the function `NOW`. Every time `NOW` is called, it will automatically return the current date and time.

Excel contains several built-in volatile functions, such as `RAND` and `TODAY`. For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](https://docs.microsoft.com/en-us/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling. For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.

To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample. Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
