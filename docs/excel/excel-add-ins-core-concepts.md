---
title: Fundamental programming concepts with the Excel JavaScript API
description: Use the Excel JavaScript API to build add-ins for Excel.
ms.date: 05/08/2019
localization_priority: Priority
---


# Fundamental programming concepts with the Excel JavaScript API

This article describes how to use the [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) to build add-ins for Excel 2016 or later. It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.

## Asynchronous nature of Excel APIs

The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office on Windows and runs inside an HTML iFrame in Office Online. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action. The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.

## Excel.run

**Excel.run** executes a function where you specify the actions to perform against the Excel object model. **Excel.run** automatically creates a request context that you can use to interact with Excel objects. When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.

The following example shows how to use **Excel.run**. The catch statement catches and logs errors that occur within the **Excel.run**.

```js
Excel.run(function (context) {
    // You can use the Excel JavaScript API here in the batch function
    // to execute actions on the Excel object model.
    console.log('Your code goes here.');
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Run options

**Excel.run** has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object. This contains a set of properties that affect platform behavior when the function runs. The following property is currently supported:

- `delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode. When **true**, the batch request is delayed and runs when the user exits cell edit mode. When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user). The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## Request context

Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.

## Proxy objects

The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.

For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object. The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### sync()

Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document. The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.

The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.

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

In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.

Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript). Doing so ensures that the **sync()** operation completes before the script continues to run. For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](/office/dev/add-ins/excel/performance).

### load()

Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it. To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load. 

> [!NOTE]
> If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method. The **load()** method is only required when you want to read properties on a proxy object.

Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.

In the following example, only specific properties of the range are loaded.

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);

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

In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.

To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md). For more information about the **load()** method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).

## null or blank property values

### null input in 2-D Array

In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.

For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### null input for a property

`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.

```js
range.values = null;
```

Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.

```js
range.format.fill.color =  null;
```

### null property values in the response

Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:

- If all cells in the range have the same font color, `range.format.font.color` specifies that color.
- If multiple font colors are present within the range, `range.format.font.color` is `null`.

### Blank input for a property

When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:

- If you specify a blank value for the `values` property of a range, the content of the range is cleared.

- If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.

- If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.

### Blank property values in the response

For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## Read or write to an unbounded range

### Read an unbounded range

An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:

- Range addresses comprised of entire column(s):<ul><li>`C:C`</li><li>`A:F`</li></ul>
- Range addresses comprised of entire row(s):<ul><li>`2:2`</li><li>`1:4`</li></ul>

When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.

### Write to an unbounded range

You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large. For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range. The API will return an error if you attempt to set cell-level properties for an unbounded range.

```js
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## Read or write to a large range

If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.

> [!IMPORTANT]
> Excel Online has a payload size limit for requests and responses of **5MB**. `RichAPI.Error` will be thrown if that limit is exceeded.

## Update all cells in a range

To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.

The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = context.workbook.worksheets.getItem(sheetName);

    var range = worksheet.getRange(rangeAddress);
    range.numberFormat = 'm/d/yyyy';
    range.values = '3/11/2015';
    range.load('text');

    return context.sync()
      .then(function () {
        console.log(range.text);
    });
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
      console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Handle errors

When an API error occurs, the API returns an **error** object that contains a code and a message. For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).

## See also

- [Get started with Excel add-ins](excel-add-ins-get-started-overview.md)
- [Excel add-ins code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
- [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md)
- [Excel JavaScript API performance optimization](/office/dev/add-ins/excel/performance)
- [Excel JavaScript API reference](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
