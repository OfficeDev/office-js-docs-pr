---
ms.date: 10/22/2025
description: Understand Excel custom functions that don't use a shared runtime and their specific JavaScript-only runtime.
title: JavaScript-only runtime for custom functions
ms.localizationpriority: medium
---

# JavaScript-only runtime for custom functions

Custom functions that don't use a shared runtime rely on a [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime). This runtime is optimized for fast calculation but has fewer APIs available.

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

This JavaScript-only runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by custom functions and the task pane (which runs in a different runtime) to store data.

## Request external data

Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

Be aware that custom functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).

A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST). Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`. You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.

## Store and access data

Within a custom function that doesn't use a shared runtime, you can store and access data by using the [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) object. The `Storage` object is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), which cannot be used by custom functions that use the JavaScript-only runtime. The `Storage` object offers 10 MB of data per domain. Domains can be shared by more than one add-in.

The `Storage` object is a shared storage solution, meaning multiple parts of an add-in are able to access the same data. For example, tokens for user authentication may be stored in the `Storage` object because it can be accessed by both a custom function (using the JavaScript-only runtime) and a task pane (using a full webview runtime). Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through the `Storage` object. Note that add-ins which have different subdomains will have different instances of `Storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).

Because the `Storage` object can be a shared location, it is important to realize that it is possible to override key-value pairs.

The following methods are available on the `Storage` object.

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

The following code sample calls the `OfficeRuntime.storage.setItem` method to set a key and value into `storage`.

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## Compare with shared runtime

Need UI integration or Office.js objects and events? Move those functions to a [shared runtime](../testing/runtimes.md#shared-runtime).

## Next steps

Learn how to [debug custom functions](custom-functions-debugging.md).

## See also

- [Authentication for custom functions without a shared runtime](custom-functions-authentication.md)
- [Create custom functions in Excel](custom-functions-overview.md)
- [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
- [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime)
