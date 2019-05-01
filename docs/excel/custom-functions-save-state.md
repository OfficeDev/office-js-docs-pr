---
ms.date: 04/25/2019
description: Use `Office.Storage` to save state with custom functions. 
title: Save and share state in custom functions (preview)
localization_priority: Priority
---

## Save and share state in custom functions

Use the `Office.Storage` object to save state related to custom functions or the task pane in your add-in. Storage is limited to 10 MB per domain (which may be shared across multiple add-ins). On Excel for Windows, the `Storage` object is a separate location within the custom functions runtime, but for Excel Online and Excel for Mac, the `Storage` object is the same as the browser's `localStorage`.

There are multiple ways to use `Office.Storage` for state management:

- You can store default values for custom functions to use when you are offline and unable to reach a web resource.
- You can save values for custom functions to use to avoid making additional calls to a web resource.
- You can save values from your custom function.
- You can store values from your task pane.

The following code sample illustrates how to store an item into `Office.Storage` and retrieve it.

```js
function storeValue(key, value) {
  return Office.Storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}

function GetValue(key) {
  return Office.Storage.getItem(key);
}

CustomFunctions.associate("STOREVALUE", StoreValue);
CustomFunctions.associate("GETVALUE", GetValue);
```

[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane.

>[!NOTE]
> The `Office.Storage` object replaces the previous storage object named `AsyncStorage` which is now deprecated. If using the `AsyncStorage` object in your current custom functions code, please update it to use the `Office.Storage` object.

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Custom functions debugging](custom-functions-debugging.md)
