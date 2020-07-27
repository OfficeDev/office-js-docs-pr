---
title: Fundamentals of the promise-based API Model
description: 'Use the promise-based API model for Excel, OneNote, and Word add-ins.'
ms.date: 07/27/2020
localization_priority: Normal
---

# Fundamentals of the promise-based API Model

This article describes how to use the API model for building add-ins in Excel, Word, and OneNote. It introduces core concepts that are fundamental to using the promise-based APIs. Note that this model is not supported by Office 2013 clients. Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions. For full platform availability notes, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).

## Asynchronous nature of the promise-based APIs

The web-based add-ins run inside a browser container. This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and run inside an HTML iFrame in Office on the web. Enabling the Office.js API to interact synchronously with the Office host across all supported platforms is not feasible due to performance considerations. Therefore, the `sync()` API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action. The following sections describe how to accomplish this using the `run()` and `sync()` APIs.

## *.run function

`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote. `*.run` automatically creates a request context that you can use to interact with Office objects. When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.

The following example shows how to use `Excel.run`. The same pattern is also used with Word and OneNote.

```js
Excel.run(function (context) {
    // You can use the Excel JavaScript API here in the batch function
    // to execute actions on the Excel object model.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Request context

The Office host and your add-in run in two different processes. Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.

## Proxy objects

The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office host and run. These APIs are fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.

For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object. The `selectedRange` object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### sync()

Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document. The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The `sync()` method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.

The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript Promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.

Because `sync()` is an asynchronous operation that returns a promise, you should always `return` the promise (in JavaScript). Doing so ensures that the `sync()` operation completes before the script continues to run. For more information about optimizing performance with `sync()`, see [Excel JavaScript API performance optimization](../excel/performance.md).

### load()

Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it. To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.

> [!NOTE]
> If you are only calling methods or setting properties on a proxy object, you do not need to call the `load()` method. The `load()` method is only required when you want to read properties on a proxy object.

Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.

In the following example, only specific properties of the `Excel.Range` object are loaded.

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load(['address', 'format/*', 'format/fill', ]);

    return context.sync()
      .then(function () {
        console.log (myRange.address);              // ok
        console.log (myRange.format.wrapText);      // ok
        console.log (myRange.format.fill.color);    // ok
        //console.log (myRange.format.font.color);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

In the previous example, because `format/font` is not specified in the call to `myRange.load()`, the `format.font.color` property cannot be read.

To optimize performance, you should explicitly specify the properties to load when using the `load()` method on an object, as covered in [Excel JavaScript API performance optimizations](../excel/performance.md).

## Handle errors

When an API error occurs, the API returns an `error` object that contains a code and a message. For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).

## See also

- [Common coding issues and unexpected platform behaviors](../develop/common-coding-issues.md).
