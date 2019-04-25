---
ms.date: 04/25/2019
description: Use `Office.Storage` to save state with custom functions. 
title: Save and share state in custom functions (preview)
localization_priority: Priority
---

## Save and share state in custom functions

The main storage location for custom functions is `Office.Storage`. Storage is limited to 10 MB per domain (which may be shared across multiple add-ins). On Excel for Windows, `Office.Storage` is a separate location within the custom functions runtime, but for Excel Online and Excel for Mac, `Office.Storage` is the same as the browser's `localStorage`. `Office.Storage` is a useful storage location that can be accessed by both custom functions and your add-in's task pane.

There are multiple ways to use `Office.Storage` to save and share state:

- You can store default values for custom functions to use when you are offline and unable to reach a web resource
- You can save values for custom functions to use to avoid making additional calls to a web resource
- You can save values here from your custom function for your task pane to use
- You can store values from your task pane here for your custom function to use

The following code sample illustrates how to storing an item into `Office.Storage` and then retrieve it.

```js
function StoreValue(key, value) {

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

[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane. Note that at this time, this sample uses `AsyncStorage` rather than `Office.Storage`, which was the previous name for this storage location. Some of the method names for `AsyncStorage` have changed, so we recommend consulting the `Office.Storage` reference documentation for correct methods.

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Custom functions debugging](custom-functions-debugging.md)
