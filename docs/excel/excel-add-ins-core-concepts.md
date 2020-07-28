---
title: Fundamental programming concepts with the Excel JavaScript API
description: 'Use the Excel JavaScript API to build add-ins for Excel.'
ms.date: 07/13/2020
localization_priority: Priority
---


# Fundamental programming concepts with the Excel JavaScript API

This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later. It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.

## Asynchronous nature of Excel APIs

The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office on Windows and runs inside an HTML iFrame in Office on the web. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the `sync()` API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action. The following sections describe how to accomplish this using the `Excel.run()` and `sync()` APIs.

## Excel.run

`Excel.run` executes a function where you specify the actions to perform against the Excel object model. `Excel.run` automatically creates a request context that you can use to interact with Excel objects. When `Excel.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.

The following example shows how to use `Excel.run`. The catch statement catches and logs errors that occur within the `Excel.run`.

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

`Excel.run` has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object. This contains a set of properties that affect platform behavior when the function runs. The following property is currently supported:

- `delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode. When **true**, the batch request is delayed and runs when the user exits cell edit mode. When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user). The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## Request context

Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a `RequestContext` object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.

## Proxy objects

The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.

For example, the following code snippet declares the local JavaScript object `selectedRange` to reference a selected range in the Excel document, and then sets some properties on that object. The `selectedRange` object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### sync()

Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Excel document. The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The `sync()` method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.

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

Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call `context.sync()`. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it. To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.

> [!NOTE]
> If you are only calling methods or setting properties on a proxy object, you do not need to call the `load()` method. The `load()` method is only required when you want to read properties on a proxy object.

Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.

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

In the previous example, because `format/font` is not specified in the call to `myRange.load()`, the `format.font.color` property cannot be read.

To optimize performance, you should explicitly specify the properties to load when using the `load()` method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md). For more information about the `load()` method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).

## null or blank property values

### null input in 2-D Array

In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.

For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### null input for a property

`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.

```js
range.values = null;
```

Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.

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

For details on the system limitations, see [Excel data transfer limits](../develop/common-coding-issues.md#excel-data-transfer-limits).

## Handle errors

When an API error occurs, the API returns an `error` object that contains a code and a message. For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).

## See also

- [Build your first Excel add-in](../quickstarts/excel-quickstart-jquery.md)
- [Excel add-ins code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md)
- [Excel JavaScript API performance optimization](../excel/performance.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
- [Common coding issues and unexpected platform behaviors](../develop/common-coding-issues.md).



---
title: Advanced programming concepts with the Excel JavaScript API
description: 'Learn how an Excel add-in interacts with objects in Excel by using the Office JavaScript API bject models.'
ms.date: 07/01/2020
localization_priority: Priority
---

# Advanced programming concepts with the Excel JavaScript API

This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.

## Office.js APIs for Excel

An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:

* **Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.

* **Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.

While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API. For example:

- [Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.

- [Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.

The following image illustrates when you might use the Excel JavaScript API or the Common APIs.

![Image of the differences between the Excel JS API and Common APIs](../images/excel-js-api-common-api.png)

## Requirement sets

Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs. To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).

### Checking for requirement set support at runtime

The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### Defining requirement set support in the manifest

You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.

The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> To make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.

### Requirement sets for the Office.js Common API

For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).

## Loading the properties of an object

Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs. The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.

### Method details

#### `load(propertyNames?: string | string[])`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Syntax

```js
object.load(param);
```

#### Parameters

|**Parameter**|**Type**|**Description**|
|:------------|:-------|:----------|
|`propertyNames`|object|Optional. Accepts property names as comma-delimited string or an array.|

#### Returns

void

#### Example

The following code sample sets the properties of one Excel range by copying the properties of another range. Note that the source object must be loaded first, before its property values can be accessed and written to the target range. This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange);
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Load option properties

As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.

|**Property**|**Type**|**Description**|
|:-----------|:-------|:----------|
|`select`|object|Contains a comma-delimited list or an array of scalar property names. Optional.|
|`expand`|object|Contains a comma-delimited list or an array of navigational property names. Optional.|
|`top`|int| Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.|
|`skip`|int|Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.|

The following code sample loads a worksheet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection. It also specifies that only the top five worksheets in the collection should be loaded. You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

### Calling `load` without parameters

If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded. To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.

> [!IMPORTANT]
> The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service. To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them. The following properties are excluded from such load operations:
>
> * `Excel.Range.numberFormatCategories`

## Scalar and navigation properties

There are two categories of properties: **scalar** and **navigational**. Scalar properties are assignable types such as strings, integers, and JSON structs. Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property. For example, `name` and `position` members on the [Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties. `prompt` on the [DataValidation](/javascript/api/excel/excel.datavalidation) object is an example of a scalar property that must be set using a JSON object (`dv.prompt = { title: "MyPrompt"}`), instead of setting the sub-properties (`dv.prompt.title = "MyPrompt" // will not set the title`).

### Scalar properties and navigation properties with `object.load()`

Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded. Additionally, navigation properties cannot be loaded directly. Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property. For example, to load the font name for a range, you must specify the `format` and `font` navigation properties as the path to the `name` property:

```js
someRange.load("format/font/name")
```

> [!NOTE]
> With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path. For example, you could set the font size for a range by using `someRange.format.font.size = 10;`. You do not need to load the property before you set it. 

## Setting properties of an object

Setting properties on an object with nested navigation properties can be cumbersome. As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API. With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.

> [!NOTE]
> The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API. The common (shared) APIs do not support this method. 

### set (properties: object, options: object)

Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.

#### Syntax

```js
object.set(properties[, options]);
```

#### Parameters

|**Parameter**|**Type**|**Description**|
|:------------|:--------|:----------|
|`properties`|object|Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.|
|`options`|object|Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)|

#### Returns

void

#### Example

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

## &#42;OrNullObject methods

Many Excel JavaScript API methods will return an exception when the condition of the API is not met. For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception. 

Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API. An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist. For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection. The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object. The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.

The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method. If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## See also

* [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
* [Excel add-ins code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API performance optimization](performance.md)
* [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
