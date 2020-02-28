---
title: Show or hide an Office Add-in in a shared runtime
description: 'Learn how to programmatically hide or show the UI of an add-in while it runs continuously'
ms.date: 03/02/2020
localization_priority: Normal
---

# Show or hide an Office Add-in in a shared runtime (preview)

An Office Add-in can include any of the following parts:

- A task pane
- A UI-less function file
- An Excel custom function

By default, each part runs in its own separate JavaScript runtime, with its own global object and global variables. 

It's possible for add-ins with two or more parts to share a common JavaScript runtime. This shared runtime feature enables new preview APIs that hide and reopen the task pane while the add-in runs.

> [!INCLUDE [Information about using preview APIs](../includes/excel-shared-runtime-preview-note.md)]

## Configure an add-in to use a shared runtime

To configure the add-in to use a shared runtime, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).

## Show and hide the task pane

The new APIs are in the `Office.addin` namespace. To show the task pane, your code calls `Office.addin.showAsTaskpane()`. Office will display in a task pane the page that you assigned to the resource ID (`resid`) for the task pane. This is the `resid` that you assigned to the `<SourceLocation>` of the `<Action xsi:type="ShowTaskpane">` in the manifest. (See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).)

This is an asynchronous method, so your code should await it when the subsequent code should not run until it completes. Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using. The following assumes that there is an Excel worksheet named **CurrentQuarterSales**. The add-in should make the task pane visible whenever this worksheet is activated. The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) event which has been registered for the worksheet.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

To hide the task pane, your code calls `Office.addin.hide()`. The following example is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) event.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### Preservation of state and event listeners

The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane. They do not unload or reload it (or reinitialize its state).

Consider the following scenario: A task pane is designed with tabs. The **Home** tab is open when the add-in is first launched. Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event. Still later code calls `showAsTaskpane()` in response to another event. The task pane will reappear, and the **Settings** tab is still selected.

![A screenshot of task pane that has four tabs labelled Home, Settings, Favorites, and Accounts.](../images/TaskpaneWithTabs.png)

In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.

Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**. The activated handler causes a green dot to appear in the task pane. The deactivated handler turns the dot red (which is its default state). Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red. While the task pane is hidden, **Sheet1** is activated. Later code calls `showAsTaskpane()` in response to some event. When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.

### Handle visibility changed event

When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the VisibilityModeChanged event. It can be useful to handle this event. For example, suppose the task pane displays a list of all the sheets in a workbook. If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list. But your code can respond to the VisibilityModeChanged event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.

To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts. Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-). The following is an example. Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

The function returns another function that *deregisters* the handler. Here is a simple, but not robust, example:

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

The `onVisibilityModeChanged` method is asynchronous which means that if your code calls the *deregister* handler that `onVisibilityModeChanged` returns, you should ensure that `onVisibilityModeChanged` has completed before calling the deregister handler. One way to do that is to use the `await` keyword on the method call as in the following example.

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

If you want to use only pre-ES2015 JavaScript, your code can use the `then` method to wait until the returned Promise object has resolved and assign the returned function to a global variable as in the following example.

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

The deregister function is itself asynchronous. So, if you have code that should not run until after the deregistration is complete, then the deregister function should also be awaited with either the `await` keyword or with a `then` method as in the following examples.

To deregister the handler:

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
