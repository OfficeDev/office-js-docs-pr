---
title: Using the host-specific API model
description: 'Learn about the promise-based API model for Excel, OneNote, and Word add-ins.'
ms.date: 07/27/2020
localization_priority: Normal
---

# Using the host-specific API model

This article describes how to use the API model for building add-ins in Excel, Word, and OneNote. It introduces core concepts that are fundamental to using the promise-based APIs. Note that this model is not supported by Office 2013 clients. Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions. For full platform availability notes, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).

> [!NOTE]
> The examples in this page use the Excel JavaScriptAPIs, but the concepts equally apply to Excel, Word, and OneNote add-ins, as well as SharePoint-embedded Visio diagrams that use the Office JavaScript APIs.

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

The Office host and your add-in run in two different processes. Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables. This `RequestContext` object is provided as an argument when calling `*.run`.

## Proxy objects

The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office host and run. These APIs are fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.

For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object. The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.

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

Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it. To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load. The following example shows the `Range.address` property being loaded for `myRange`.

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
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

> [!NOTE]
> If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method. The `load()` method is only required when you want to read properties on a proxy object.

Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.

#### Scalar and navigation properties

There are two categories of properties: **scalar** and **navigational**. Scalar properties are assignable types such as strings, integers, and JSON structs. Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property. For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.

Your add-in can use navigational properties as a path to load specific scalar properties. The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.

```js
someRange.load("format/font/name")
```

You can also set the scalar properties of a navigation property by traversing the path. For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`. You don't need to load the property before you set it.

#### Calling `load` without parameters

If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded. To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.

> [!IMPORTANT]
> The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service. To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them. The following properties are excluded from such load operations:
>
> * `Excel.Range.numberFormatCategories`

To optimize performance, you should always explicitly specify the properties to load when using the `load()` method on an object.

### ClientResult

Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm. As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection. `getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number. Your script can't access that value until `context.sync()` is called.

The following code gets the total number of tables in an Excel workbook and logs that number to the console.

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### set()

Setting properties on an object with nested navigation properties can be cumbersome. As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs. With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.

The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## get&#42;OrNullObject methods

Some `get*` methods throw an exception when the desired object is not present. For example, if you attempt to get an Excel worksheet by specifying a worksheet name that is not in the workbook, the `getItem()` method throws an `ItemNotFound` exception.

Any `get*OrNullObject` method variant lets you check for an object without throwing exceptions. These methods return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist. For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection. The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object. The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.

The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method. If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet.
    }

    dataSheet.position = 1;
    //...
  })
```

## See also

* [Common coding issues and unexpected platform behaviors](../develop/common-coding-issues.md).
