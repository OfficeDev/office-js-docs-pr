---
title: Understanding the JavaScript API for Office
description: ''
ms.date: 01/23/2018
---


# Understanding the JavaScript API for Office

This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)). 

## Referencing the JavaScript API for Office library in your add-in

The [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.

For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## Initializing your add-in

**Applies to:** All add-in types

Your code must not call any Office.js APIs until the library is fully loaded. There are two events that your code can handle to ensure that the library is loaded. They are described in the sections below. We recommend that you use the newer, more flexible, `Office.onReady`. But the older `Office.initialize` is still supported. The main differences between them:

- You can assign only one handler to `Office.initialize`. But you can assign multiple handlers, in different places in your code, to `Office.onReady`.

- The `Office.initialize` event fires only once: at the end of the internal process in which Office.js initializes itself. And it fires immediately after the internal process ends. If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run. As a practical matter, this means that you must assign the handler as virtually the first thing your custom script does. The `Office.onReady` event fires when `Office.initialize` does but it also fires, in effect, whenever another handler is assigned to it. For example, you could have a handler with custom initialization logic assigned to `Office.onReady` as soon as your custom script loads (in, for example, an Immediately Invoked Function Expression). But you could also have a button in the task pane, whose script assigns another handler to `Office.onReady`. If so, the second handler runs when the button is clicked.

- The `Office.onReady` event has a default handler that returns a Promise object. This gives you an alternative way of responding to the event. See **Initialize with Office.onReady** below.

- Your code *must* assign some handler to `Office.initialize` although it can be a function that does nothing. See **Initialize with Office.initialize** below. You are not required to assign anything to `Office.onReady`.

For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).

### Initialize with Office.onReady

`Office.onReady` fires first when the Office.js library is fully loaded and ready for user interaction. You can use the `Office.onReady` event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values. You can also use the handler to initialize other custom logic for your add-in, such as establishing bindings, prompting for default add-in settings values, and so on. For example, the following handler checks to see that the user's version of Excel supports all the APIs that the add-in might call.

```js
Office.onReady = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be placed within the `Office.onReady` event. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:

```js
Office.onReady = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

If you prefer, you can call the default handler for `Office.onReady` and respond to the event when the Promise is resolved. This is the syntax in JavaScript:

```js
Office.onReady()
    .then(function() {
        console.log("Office is now ready");
    });
```

Here is the same example using TypeScript:

```ts
(async () => {
    await Office.onReady();
    console.log("Office is now ready!");
})();
```

> [!NOTE]
> Even if you use `Office.onReady`, you must still assign some function to  `Office.initialize`, even if it is a function that does nothing. See **Initialize with Office.initialize** below. 

### Initialize with Office.initialize

`Office.initialize` fires when the Office.js library is fully loaded and ready for user interaction. You can use the `Office.initialize` event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values. You can also use the handler to initialize other custom logic for your add-in, such as establishing bindings, prompting for default add-in settings values, and so on.

At a minimum, the initialize event assignment would look like the following example, in which a function that does nothing is assigned:

```js
Office.initialize = function () { };
```

If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be placed within the `Office.initialize` event. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

All pages within an Office Add-ins are required to assign an event handler to the `Office.initialize` event.
If you fail to assign an event handler, your add-in may raise an error when it starts. Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run. If you don't need any initialization code, then the body of the function you assign to `Office.initialize` can be empty, as it is in the first example above.

#### Initialization reason

For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter. This parameter can be used to determine how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.

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

For more information, see [Office.initialize Event](https://dev.office.com/reference/add-ins/shared/office.initialize) and [InitializationReason Enumeration](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration). 

## Office JavaScript API object model

Once initialized, the add-in can interact with the host (e.g. Excel, Outlook). The [Office JavaScript API object model](/office-javascript-api-object-model.md)) page has more details on specific usage patterns. There is also detailed reference documentation for both [shared APIs](https://dev.office.com/reference/add-ins/javascript-api-for-office) and specific hosts.

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
