---
ms.date: 05/13/2020
description: 'Use `OfficeRuntime.storage` to save state and data with UI-less custom functions.' 
title: Handling data and state in UI-less custom functions
localization_priority: Normal
---

# Handling data and state in UI-less custom functions

UI-less custom functions don't run in the shared runtime like ordinary custom functions. In the UI-less custom functions runtime, only calculations and web requests can be performed and interaction with the task pane requires using a separate object. The UI-less custom functions runtime doesn't have access to the Document Object Model (DOM), so it can't use jQuery or other libraries which rely on the DOM.

Instead UI-less custom functions rely on a `storage` object to interact with the task pane. UI-less custom functions running on Excel on Windows have access to the `OfficeRuntime.storage` object, a separate location within the UI-less custom functions runtime. For UI-less functions running on Excel on the web and Mac, the `storage` object is the same as the browser's `localStorage`.

Use the `storage` object to save state when using UI-less custom functions. Storage is limited to 10 MB per domain, may be shared across multiple add-ins, and is unencrypted. If two add-ins share a domain (for example, www.contoso.com/addin1, www.contoso.com/addin2), they can also share information through `storage`. However, if add-ins have different subdomains, (such as subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2), they will have different `storage` instances. Because `storage` can be a shared location, be aware that you can overwrite key-value pairs.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

There are multiple ways to use `storage` for state management:

- You can store default values for UI-less custom functions to use when you are offline and unable to reach a web resource.
- You can save values for UI-less custom functions to use to avoid making additional calls to a web resource.
- You can save values from your UI-less custom function.
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

The following methods are available on the `storage` object:

- getItem
- getItems
- setItem
- setItems
- removeItem
- removeItems
- getKeys

> [!NOTE]
There's no method for clearing all information (such as `clear`). Instead, use `removeItems` to remove multiple entries at a time.

## Addressing cell's context parameter

In some cases you need to get the address of the cell that invoked your UI-less custom function. This is useful in the following scenarios:

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

## Requesting external data

Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).

Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST). Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`. You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.

## Next steps
Learn how to [autogenerate the JSON metadata for your custom functions](custom-functions-json-autogeneration.md).

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Ui-less custom functions debugging](custom-functions-debugging.md)
