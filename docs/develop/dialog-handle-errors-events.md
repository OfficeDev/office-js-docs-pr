---
title: Handling errors and events in the Office dialog box 
description: Learn how to trap and handle errors when opening and using the Office dialog box.
ms.date: 03/11/2025
ms.topic: error-reference
ms.localizationpriority: medium
---

# Handle errors and events in the Office dialog box

This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.

> [!NOTE]
> This article presupposes that you're familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).
>
> See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).

Your code should handle two categories of events.

- Errors returned by the call of `displayDialogAsync` because the dialog box can't be created.
- Errors, and other events, in the dialog box.

## Errors from displayDialogAsync

In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.

|Code number|Meaning|
|:-----|:-----|
|12004|The domain of the URL passed to `displayDialogAsync` isn't trusted. The domain must be the same domain as the host page (including protocol and port number).<br><br>In Outlook on the web and the [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), this error occurs when an add-in is hosted on a localhost server and its manifest doesn't specify an [AppDomain](/javascript/api/manifest/appdomain) element for localhost.|
|12005|The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.|
|12009|The user chose to ignore the dialog box. This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box. For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handle-pop-up-blockers-with-office-on-the-web).|
|12011| The add-in is running in Office on the web and the user's browser configuration is blocking popups. This most commonly happens when the browser is Edge Legacy (an older, unsupported webview) and the domain of the add-in is in different security zone from the domain that the dialog is trying to open. Another scenario which triggers this error is that the browser is Safari and it's configured to block all popups. Consider responding to this error with a prompt to the user to change their browser configuration or use a different browser.|

When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function. When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object. For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page). When the call to `displayDialogAsync` fails, the dialog box isn't created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated. You should always provide a callback that tests the `status` and responds when it's an error. For an example that reports the error message regardless of its code number, see the following code. (The `showNotification` function, not defined in this article, either displays or logs the error. For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)

```js
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## Errors and events in the dialog box

Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page. For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).

|Code number|Meaning|
|:-----|:-----|
|12002|One of the following:<ul><li>No page exists at the URL that was passed to `displayDialogAsync`.</li><li>The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it can't find or load, or it has been directed to a URL with invalid syntax.</li></ul>|
|12003|The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.|
|12006|One of the following:<ul><li>The dialog box was closed, usually because the user chose the **Close** button **X**.</li><li>The dialog returned a [Cross-Origin-Opener-Policy: same-origin](https://developer.mozilla.org/docs/Web/HTTP/Headers/Cross-Origin-Opener-Policy) response header. To prevent this, you must set the header to `Cross-Origin-Opener-Policy: unsafe-none` or configure your add-in and dialog to be in the same domain as the host page.</li></ul>|

Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example.

```js
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example.

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it can't find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

## See also

For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).
