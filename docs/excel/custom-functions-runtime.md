---
ms.date: 09/25/2020
description: Understand Excel custom functions that don't use a task pane and their specific JavaScript runtime.
title: Runtime for UI-less Excel custom functions
localization_priority: Normal
---

# Runtime for UI-less Excel custom functions

Custom functions that don't use a task pane (UI-less custom functions) use a JavaScript runtime that is designed to optimize performance of calculations.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

This JavaScript runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by UI-less custom functions and the task pane to store data.

## Requesting external data

Within a UI-less custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

Be aware that UI-less functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).

A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST). Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`. You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.

## Storing and accessing data

Within a UI-less custom function, you can store and access data by using the `OfficeRuntime.storage` object. `Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), which cannot be used by UI-less custom functions. `Storage` offers 10 MB of data per domain. Domains can be shared by more than one add-in.

`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data. For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a UI-less custom function and add-in UI elements such as a task pane. Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`. Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).

Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.

The following methods are available on the `storage` object.

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> There's no method for clearing all information (such as `clear`). Instead, you should instead use `removeItems` to remove multiple entries at a time.

### OfficeRuntime.storage example

The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## Additional considerations

If your add-in only uses UI-less custom functions, note that you can't access the Document Object Model (DOM) with UI-less custom functions or use libraries like jQuery that rely on the DOM.

## Next steps
Learn how to [debug UI-less custom functions](custom-functions-debugging.md).

## See also

* [Authenticate UI-less custom functions](custom-functions-authentication.md)
* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
