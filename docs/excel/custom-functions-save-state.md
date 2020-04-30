---
ms.date: 04/29/2019
description: 'Use `OfficeRuntime.storage` to save state with custom functions.' 
title: Save and share state in UI-less custom functions
localization_priority: Normal
---

# Save and share state in custom functions

Use the `OfficeRuntime.storage` object to save state related to custom functions or the task pane in your add-in. Storage is limited to 10 MB per domain (which may be shared across multiple add-ins). In Excel on Windows, the `storage` object is a separate location within the custom functions runtime, but for Excel on the web and Mac, the `storage` object is the same as the browser's `localStorage`.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

There are multiple ways to use `storage` for state management:

- You can store default values for custom functions to use when you are offline and unable to reach a web resource.
- You can save values for custom functions to use to avoid making additional calls to a web resource.
- You can save values from your custom function.
- You can store values from your task pane.

The following code sample illustrates how to store an item into `storage` and retrieve it.

```js
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}
```

[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane.

>[!NOTE]
> The `storage` object replaces the previous storage object named `AsyncStorage` which is now deprecated. If using the `AsyncStorage` object in your current custom functions code, please update it to use the `storage` object.

## Addressing cell's context parameter

In some cases you need to get the address of the cell that invoked your custom function. This is useful in the following scenarios:

- Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](../excel/custom-functions-runtime.md#storing-and-accessing-data). Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.
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
```

By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`. For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.

## Next steps
Learn how to [autogenerate the JSON metadata for your custom functions](custom-functions-json-autogeneration.md). 

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Custom functions debugging](custom-functions-debugging.md)
