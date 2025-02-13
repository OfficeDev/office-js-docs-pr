---
title: Using the application-specific API model
description: Learn about the promise-based API model for Excel, OneNote, PowerPoint, Visio, and Word add-ins.
ms.date: 02/12/2025
ms.localizationpriority: medium
---

# Application-specific API model

This article describes how to use the API model for building add-ins in Excel, OneNote, PowerPoint, Visio, and Word. It introduces core concepts that are fundamental to using the promise-based APIs.

> [!NOTE]
> This model isn't supported by Outlook or Project clients. Use the [Common API model](office-javascript-api-object-model.md) to work with those applications. For full platform availability notes, see [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets).

> [!TIP]
> The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, PowerPoint, Visio, and Word JavaScript APIs. For complete code samples that show how you could use these and other concepts in various Office applications, see [Office Add-in code samples](../overview/office-add-in-code-samples.md).

## Asynchronous nature of the promise-based APIs

Office Add-ins are websites which appear inside a webview control within Office applications, such as Excel. This control is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iframe in Office on the web. Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms. Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action. The following sections describe how to accomplish this using the `run()` and `sync()` APIs.

## *.run function

`Excel.run`, `OneNote.run`, `PowerPoint.run`, and `Word.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote. `*.run` automatically creates a request context that you can use to interact with Office objects. When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.

The following example shows how to use `Excel.run`. The same pattern is also used with OneNote, PowerPoint, Visio, and Word.

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
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

The Office application and your add-in run in different processes. Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables. This `RequestContext` object is provided as an argument when calling `*.run`.

## Proxy objects

The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run. These APIs are fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.

For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object. The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.

```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### Performance tip: Minimize the number of proxy objects created

Avoid repeatedly creating the same proxy object. Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
const range = worksheet.getRange("A1");
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties
// without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

### sync()

Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document. The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.

The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.

```js
await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    await context.sync();
    console.log('The selected range is: ' + selectedRange.address);
});
```

In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.

Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run. If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.

#### Performance tip: Minimize the number of sync calls

In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web. To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it. For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).

### load()

Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it. To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load. The following example shows the `Range.address` property being loaded for `myRange`.

```js
await Excel.run(async (context) => {
    const sheetName = 'Sheet1';
    const rangeAddress = 'A1:B2';
    const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');
    await context.sync();
      
    console.log (myRange.address);   // ok
    //console.log (myRange.values);  // not ok as it was not loaded

    console.log('done');
});
```

> [!NOTE]
> If you're only calling methods or setting properties on a proxy object, you don't need to call the `load()` method. The `load()` method is only required when you want to read properties on a proxy object.

Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.

#### Scalar and navigation properties

There are two categories of properties: **scalar** and **navigational**. Scalar properties are assignable types such as strings, integers, and JSON structs. Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property. For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.

Your add-in can use navigational properties as a path to load specific scalar properties. The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.

```js
someRange.load("format/font/name")
```

You can also set the scalar properties of a navigation property by traversing the path. For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`. You don't need to load the property before you set it.

Please be aware that some of the properties under an object may have the same name as another object. For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well. So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement). To avoid this, your code should only load the "leaf nodes" in an object tree.

#### Load from a collection

When working with a collection, use `load` on the collection to load properties for every object in the collection. Use `load` exactly as you would for an individual object in that collection.

The following sample code shows the `name` property being loaded and logged for every chart in the "Sample" worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chartCollection = sheet.charts;

    // Load the name property on every chart in the chart collection.
    chartCollection.load("name");
    await context.sync();

    chartCollection.items.forEach((chart) => {
        console.log(chart.name);
    });
});
```

You normally don't include the `items` property of the collection in the `load` arguments. All the items are loaded if you load any item properties. However, if you will be looping over the items in the collection, but don't need to load any particular property of the items, you need to `load` the `items` property.

The following sample code shows the `name` property being set for every chart in the "Sample" worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const chartCollection = sheet.charts;

    // Load the items property from the chart collection to set properties on individual charts.
    chartCollection.load("items");
    await context.sync();

    chartCollection.items.forEach((chart, index) => {
        chart.name = `Sample chart ${index}`;
    });
});
```

#### Calling `load` without parameters (not recommended)

If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded. Loading unneeded data will slow down your add-in. You should always explicitly specify which properties to load.

> [!IMPORTANT]
> The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service. To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them. The following properties are excluded from such load operations.
>
> - `Excel.Range.numberFormatCategories`

### ClientResult

Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm. As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection. `getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number. Your script can't access that value until `context.sync()` is called.

The following code gets the total number of tables in an Excel workbook and logs that number to the console.

```js
const tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
await context.sync();

// Trying to log the value before calling sync would throw an error.
console.log (tableCount.value);
```

### set()

Setting properties on an object with nested navigation properties can be cumbersome. As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs. With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.

The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:E2");
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

    await context.sync();
});
```

### Some properties cannot be set directly

Some properties cannot be set, despite being writable. These properties are part of a parent property that must be set as a single object. This is because that parent property relies on the subproperties having specific, logical relationships. These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties. One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout). The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here.

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`. That statement throws an error because `zoom` is not loaded. Even if `zoom` were to be loaded, the set of scale will not take effect. All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.

This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#excel-excel-range-format-member). Properties of `format` can be set using object navigation, as shown here.

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

You can identify a property that cannot have its subproperties directly set by checking its read-only modifier. All read-only properties can have their non-read-only subproperties directly set. Writeable properties like `PageLayout.zoom` must be set with an object at that level. In summary:

- Read-only property: Subproperties can be set through navigation.
- Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).

## &#42;OrNullObject methods and properties

Some accessor methods and properties throw an exception when the desired object doesn't exist. For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception. The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code. This is accomplished by using the `*OrNullObject` variations of methods and properties. These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.

For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection. The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`. Your code can then evaluate this property to determine whether the object exists.

> [!NOTE]
> The `*OrNullObject` variations do not ever return the JavaScript value `null`. They return ordinary Office proxy objects. If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`. Do not test the returned object for nullity or falsity. It is never `null`, `false`, or `undefined`.

The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method. If a worksheet with that name does not exist, a new sheet is created. Note that the code does not load the `isNullObject` property. Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `dataSheet.load('isNullObject')`.

```js
await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
    
    await context.sync();
    
    if (dataSheet.isNullObject) {
        dataSheet = context.workbook.worksheets.add("Data");
    }
    
    // Set `dataSheet` to be the second worksheet in the workbook.
    dataSheet.position = 1;
});
```

## Undo support

Undo is partially supported by the application-specific Office JavaScript APIs. This means that users may be able to revert changes made by add-ins through the undo command. If a particular API doesn't support undo, the application's undo stack is cleared. This means that you won't be able to undo the effects of that API or anything prior to calling that API.

API support for undo is continuing to expand but is currently incomplete. We advise against designing your add-in in such a way that it relies on undo support.

## See also

- [Common JavaScript API object model](office-javascript-api-object-model.md)
- [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md)
