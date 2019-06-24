---
title: Understanding the JavaScript API for Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
---

# Understanding the JavaScript API for Office

This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)). 

## Referencing the JavaScript API for Office library in your add-in

The [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.

For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## Initializing your add-in

**Applies to:** All add-in types

Office Add-ins often have start-up logic to do things such as:

- Check that the user's version of Office will support all the Office APIs that your code calls.

- Ensure the existence of certain artifacts, such as worksheet with a specific name.

- Prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.

- Establish bindings.

- Use the Office dialog API to prompt the user for default add-in settings values.

But your start-up code must not call any Office.js APIs until the library is loaded. There are two ways that your code can ensure that the library is loaded. They are described in the following sections: 

- [Initialize with Office.onReady()](#initialize-with-officeonready)
- [Initialize with Office.initialize](#initialize-with-officeinitialize)

> [!TIP]
> We recommend that you use `Office.onReady()` instead of `Office.initialize`. Although `Office.initialize` is still supported, using `Office.onReady()` provides more flexibility. You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure, but you can call `Office.onReady()` in different places in your code and use different callbacks.
> 
> For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).

### Initialize with Office.onReady()

`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is loaded. When the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.). The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.

One way to call `Office.onReady()` is to pass it a callback method. Here's an example:

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback. For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

Here is the same example using the `async` and `await` keywords in TypeScript:

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

However, there are exceptions to this practice. For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools. Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`. Another exception: you want a progress indicator to appear in the task pane while the add-in is loading. In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator. Then the Office `onReady`'s callback can replace the progress indicator with the final UI. 

### Initialize with Office.initialize

An initialize event fires when the Office.js library is loaded and ready for user interaction. You can assign a handler to `Office.initialize` that implements your initialization logic. The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event. (But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter. This parameter specifies how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).

### Major differences between Office.initialize and Office.onReady

- You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks. For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback. If so, the second callback runs when the button is clicked.

- The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself. And it fires *immediately* after the internal process ends. If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run. For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript. By the time your script loads and assigns the handler, the initialize event has already happened. But it is never "too late" to call `Office.onReady()`. If the initialize event has already happened, the callback runs immediately.

> [!NOTE]
> Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads. Some Office host and platform combinations won't load the task pane until one of these happens. The following examples show these two approaches.
>
>```js	
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## Office JavaScript API object model

Once initialized, the add-in can interact with the host (e.g. Excel, Outlook). The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns. There is also detailed reference documentation for both [Common APIs](/office/dev/add-ins/reference/javascript-api-for-office) and host-specific APIs.

## API support matrix

This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Host name**|Database|Workbook|Mailbox|Presentation|Document|Project|
||**Supported** **Host applications**|Access web apps|Excel,<br/>Excel Online|Outlook,<br/>Outlook Web App,<br/>OWA for Devices|PowerPoint,<br/>PowerPoint Online|Word|Project|
|**Supported add-in types**|Content|Y|Y||Y|||
||Task pane||Y||Y|Y|Y|
||Outlook|||Y||||
|**Supported API features**|Read/Write Text||Y||Y|Y|Y<br/>(Read only)|
||Read/Write Matrix||Y|||Y||
||Read/Write Table||Y|||Y||
||Read/Write HTML|||||Y||
||Read/Write<br/>Office Open XML|||||Y||
||Read task, resource, view, and field properties||||||Y|
||Selection changed events||Y|||Y||
||Get whole document||||Y|Y||
||Bindings and binding events|Y<br/>(Only full and partial table bindings)|Y|||Y||
||Read/Write Custom XML Parts|||||Y||
||Persist add-in state data (settings)|Y<br/>(Per host add-in)|Y<br/>(Per document)|Y<br/>(Per mailbox)|Y<br/>(Per document)|Y<br/>(Per document)||
||Settings changed events|Y|Y||Y|Y||
||Get active view mode<br/>and view changed events||||Y|||
||Navigate to locations<br/>in the document||Y||Y|Y||
||Activate contextually<br/>using rules and RegEx|||Y||||
||Read Item properties|||Y||||
||Read User profile|||Y||||
||Get attachments|||Y||||
||Get User identity token|||Y||||
||Call Exchange Web Services|||Y||||
