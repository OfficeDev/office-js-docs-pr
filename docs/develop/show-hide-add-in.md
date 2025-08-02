---
title: Show or hide the task pane of your Office Add-in
description: Learn how to programmatically hide or show the user interface of an add-in while it runs continuously.
ms.date: 02/12/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Show or hide the task pane of your Office Add-in

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

You can show the task pane of your Office Add-in by calling the `Office.addin.showAsTaskpane()` method.

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

The previous code assumes a scenario where there is an Excel worksheet named **CurrentQuarterSales**. The add-in will make the task pane visible whenever this worksheet is activated. The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member) event which has been registered for the worksheet.

You can also hide the task pane by calling the `Office.addin.hide()` method.

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

The previous code is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member) event.

## Additional details on showing the task pane

When you call `Office.addin.showAsTaskpane()`, Office will display in a task pane the file that you specified in the manifest. The configuration depends on what type of manifest you're using. 

- **Unified manifest for Microsoft 365**: The URL of the file is assigned as the value of a "runtimes.code.page" property of the runtime object which has an action object of type "openPage".

   [!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]
   
- **Add-in only manifest**: The URL of the file is assigned as the resource ID (`resid`) value of the task pane. This `resid` value can be assigned or changed by opening your manifest file and locating `<SourceLocation>` inside the `<Action xsi:type="ShowTaskpane">` element.

(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md) for additional details.)

Since `Office.addin.showAsTaskpane()` is an asynchronous method, your code will continue running until the method is complete. Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.

## Configure your add-in to use the shared runtime

To use the `showAsTaskpane()` and `hide()` methods, your add-in must use the [shared runtime](../testing/runtimes.md#shared-runtime). For more information, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).

## Preservation of state and event listeners

The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane. They do not unload or reload it (or reinitialize its state).

Consider the following scenario: A task pane is designed with tabs. The **Home** tab is open when the add-in is first launched. Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event. Still later code calls `showAsTaskpane()` in response to another event. The task pane will reappear, and the **Settings** tab is still selected.

![A task pane that has four tabs labelled Home, Settings, Favorites, and Accounts.](../images/TaskpaneWithTabs.png)

In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.

Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**. The activated handler causes a green dot to appear in the task pane. The deactivated handler turns the dot red (which is its default state). Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red. While the task pane is hidden, **Sheet1** is activated. Later code calls `showAsTaskpane()` in response to some event. When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.

## Handle the visibility changed event

When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event. It can be useful to handle this event. For example, suppose the task pane displays a list of all the sheets in a workbook. If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list. But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member) collection as shown in the example code below.

To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts. Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1)). The following is an example. Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode == "Taskpane") {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

The function returns another function that *deregisters* the handler. Here is a simple, but not robust, example.

```javascript
const removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode == "Taskpane") {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

The `onVisibilityModeChanged` method is asynchronous and returns a promise, which means that your code needs to await the fulfillment of the promise before it can call the **deregister** handler.

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
const removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode == "Taskpane") {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

The deregister function is also asynchronous and returns a promise. So, if you have code that should not run until after the deregistration is complete, then you should await the promise returned by the deregister function.

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## See also

- [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md)
- [Run code in your Office Add-in when the document opens](run-code-on-document-open.md)
