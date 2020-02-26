---
title: Show or hide an Office Add-in in a shared runtime
description: 'Learn how to programmatically hide or show the UI of an add-in while it runs continuously'
ms.date: 03/02/2020
localization_priority: Normal
---

# Show or hide an Office Add-in in a shared runtime (preview)

Office Add-ins can include any of the following parts:

- a task pane
- a UI-less function file
- an Excel custom function

Until now, if an add-in had more than one of these parts, then each part ran in its own separate JavaScript runtime, with its own global object and global variables.

Microsoft is now making it possible for add-ins with two or more parts to share a common JavaScript runtime. This new feature enables new preview APIs for hiding the task pane part of an add-in while the add-in remains running and for reopening the task pane later.

> [!IMPORTANT]
> The features described in this article are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you will need to [join Office Insider](https://insider.office.com/join).
> A good way to try out preview features is by using an Office 365 subscription. If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Configure an add-in to use a shared runtime

To configure the add-in to use a shared runtime, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).

## Show and hide the task pane

The new APIs are in the `Office.addin` namespace. To show the task pane, your code calls `Office.addin.showAsTaskpane()`. Office will display in a task pane the page that you assigned to the resource ID (`resid`) for the task pane. This is the `resid` that you assigned to the `<SourceLocation>` of the `<Action xsi:type="ShowTaskpane">` in the manifest. (See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).)

This is an asynchronous method, so code that should not run until it completes should be awaited, either with the `await` keyword or with a `then()` method, depending on which JavaScript syntax you are using. The following assumes that there is an Excel worksheet named **CurrentQuarterSales**. The task pane should be made visible whenever this worksheet is activated. The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) event which has been registered for the worksheet.

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

Consider the following scenario: A task pane is designed with tabs. The **Home** tab is open when the add-in is first launched. Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event. Still later code calls `showAsTaskpane()` in response to another event. The task pane will reappear, and the **Settings** tab is still elected.

In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.

Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Excel.Worksheet.onDeactivated` events for a sheet named **Sheet1**. The activated handler causes a green dot to appear in the task pane. The deactivated handler turns the dot red (which is its default state). Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red. While the task pane is hidden, **Sheet1** is activated. Later code calls `showAsTaskpane()` in response to some event. When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.

### Handle visibility changed event

Changing the visibility of the task pane is an event that your code can handle. To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts. Instead, there is a special function to which you pass your handler: `Office.addin.onVisibilityModeChanged`. The function returns another function that deregisters the handler. The following is an example. Note that the `visibilityMode` property of the `args` object that Office passes to the handler can have the values "Hidden" or "Taskpane".

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        console.log('Visibility is changed to ' + args.visibilityMode);
});
```

To deregister the handler:

```javascript
await removeVisibilityModeHandler();
```
