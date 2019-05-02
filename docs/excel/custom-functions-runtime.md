---
ms.date: 05/02/2019
description: Understand key scenarios in developing Excel custom functions that use the new JavaScript runtime.
title: Runtime for Excel custom functions
localization_priority: Normal
---

# Runtime for Excel custom functions

Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements. This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.

The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box. This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.

## Requesting external data

Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).

Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST). Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`. You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.

### XHR example

In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID. The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.

> [!NOTE] 
> When using fetch or XHR, a new JavaScript `Promise` is returned. Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ 
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
        };
        
        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## Receiving data via WebSockets

Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server. By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.

### WebSockets example

The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.

```JavaScript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## Storing and accessing data

Within a custom function (or within any other part of an add-in), you can store and access data by using the `Office.storage` object. `Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions. `Storage` offers 10 MB of data per domain. Domains can be shared by more than one add-in.

`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data. For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a custom function and add-in UI elements such as a task pane. Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `storage`. Note that add-ins which have different subdomains will have different instances of `storage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).

Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.

The following methods are available on the `storage` object:

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

.[!NOTE]
> There's no method for clearing all information (such as `clear`). Instead, you should instead use `removeItems` to remove multiple entries at a time.

### Office.storage example

The following code sample calls the `Office.storage.setItem` function to set a key and value into `storage`.

```JavaScript
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## Additional considerations

In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM. On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.

## Next steps
Learn how custom functions can [perform web requests](custom-functions-web-reqs.md).

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions architecture](custom-functions-architecture.md)
* [Display a dialog in custom functions](custom-functions-dialog.md)
* [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
