---
title: Initialize your Office Add-in
description: Learn how to initialize your Office Add-in.
ms.date: 08/18/2023
ms.localizationpriority: medium
---

# Initialize your Office Add-in

Office Add-ins often have start-up logic to do things such as:

- Check that the user's version of Office supports all the Office APIs that your code calls.

- Ensure the existence of certain artifacts, such as a worksheet with a specific name.

- Prompt the user to select some cells in Excel, and then insert a chart initialized with those selected values.

- Establish bindings.

- Use the Office Dialog API to prompt the user for default add-in settings values.

However, an Office Add-in can't successfully call any Office JavaScript APIs until the library has been loaded. This article describes the two ways your code can ensure that the library has been loaded.

- Initialize with `Office.onReady()`.
- Initialize with `Office.initialize`.

> [!TIP]
> We recommend that you use `Office.onReady()` instead of `Office.initialize`. Although `Office.initialize` is still supported, `Office.onReady()` provides more flexibility. You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure. You can call `Office.onReady()` in different places in your code and use different callbacks.
>
> For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).

## Initialize with Office.onReady()

`Office.onReady()` is an asynchronous function that returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) object while it checks to see if the Office.js library is loaded. When the library is loaded, it resolves the Promise as an object that specifies the Office client application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.). The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.

One way to call `Office.onReady()` is to pass it a callback function. Here's an example.

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

Here's the same example using the `async` and `await` keywords in TypeScript.

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

If you're using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`. For example, [JQuery's](https://jquery.com) `$(document).ready()` method would be referenced as follows:

```js
Office.onReady(function() {
    // Office is ready.
    $(document).ready(function () {
        // The document is ready.
    });
});
```

However, there are exceptions to this practice. For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office application) in order to debug your UI with browser tools. In this scenario, once Office.js determines that it is running outside of an Office host application, it will call the callback and resolve the promise with `null` for both the host and platform.

Another exception would be if you want a progress indicator to appear in the task pane while the add-in is loading. In this scenario, your code should call the jQuery `ready` and use its callback to render the progress indicator. Then the `Office.onReady` callback can replace the progress indicator with the final UI.

## Initialize with Office.initialize

An initialize event fires when the Office.js library is loaded and ready for user interaction. You can assign a handler to `Office.initialize` that implements your initialization logic. The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

If you're using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event (the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also). For example, [JQuery's](https://jquery.com) `$(document).ready()` method would be referenced as follows:

```js
Office.initialize = function () {
    // Office is ready.
    $(document).ready(function () {
        // The document is ready.
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

## Major differences between Office.initialize and Office.onReady

- You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks. For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback. If so, the second callback runs when the button is clicked.

- The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself. And it fires *immediately* after the internal process ends. If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run. For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript. By the time your script loads and assigns the handler, the initialize event has already happened. But it's never "too late" to call `Office.onReady()`. If the initialize event has already happened, the callback runs immediately.

> [!NOTE]
> Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads. Some Office application and platform combinations won't load the task pane until one of these happens. The following examples show these two approaches.
>
>```js
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## Debug initialization

For information about debugging the `Office.initialize` and `Office.onReady()` functions, see [Debug the initialize and onReady functions](../testing/debug-initialize-onready.md).

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md)