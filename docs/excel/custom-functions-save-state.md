---
ms.date: 05/03/2019
description: Use `Office.storage` to save state with custom functions. 
title: Save and share state in custom functions
localization_priority: Priority
---

## Save and share state in custom functions

Use the `Office.storage` object to save state related to custom functions or the task pane in your add-in. Storage is limited to 10 MB per domain (which may be shared across multiple add-ins). On Excel for Windows, the `storage` object is a separate location within the custom functions runtime, but for Excel Online and Excel for Mac, the `storage` object is the same as the browser's `localStorage`.

There are multiple ways to use `storage` for state management:

- You can store default values for custom functions to use when you are offline and unable to reach a web resource.
- You can save values for custom functions to use to avoid making additional calls to a web resource.
- You can save values from your custom function.
- You can store values from your task pane.

The following code sample illustrates how to store an item into `storage` and retrieve it.

```js
function storeValue(key, value) {
  return Office.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return Office.storage.getItem(key);
}

CustomFunctions.associate("STOREVALUE", StoreValue);
CustomFunctions.associate("GETVALUE", GetValue);
```

[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane.

>[!NOTE]
> The `storage` object replaces the previous storage object named `AsyncStorage` which is now deprecated. If using the `AsyncStorage` object in your current custom functions code, please update it to use the `storage` object.

## Next steps
Learn how to [autogenerate the JSON metadata for your custom functions](custom-functions-json-autogeneration.md). 

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Custom functions debugging](custom-functions-debugging.md)
