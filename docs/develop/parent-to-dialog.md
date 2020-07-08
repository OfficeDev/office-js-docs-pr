---
title: Passing data and messages to a dialog box from its host page
description: 'Learn how to pass data to a dialog from the host page by using the messageChild and DialogParentMessageReceived APIs.'
ms.date: 07/07/2020
localization_priority: Normal
---

# Passing data and messages to a dialog box from its host page (preview)

Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using the [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method of the [Dialog](/javascript/api/office/office.dialog) object.

> [!Important]
>
> - The APIs described in this article are in preview. They are available to developers for experimentation; but should not be used in a production add-in. Until this API is released, use the techniques described in [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) for production add-ins.
> - The APIs described in this article require a Microsoft 365 subscription. You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://insider.office.com). Please note that when a build graduates to the production semi-annual channel, support for preview features is turned off for that build.
> - In the initial stage of the preview, the APIs are supported in Excel, PowerPoint, and Word; but not in Outlook.
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## Use `messageChild()` from the host page

When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned. It should be assigned to a variable, which typically has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-)
method because the object will be referenced by other methods. The following is an example:

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, or stringified data, to the dialog box. This raises a `DialogParentMessageReceived` event in the dialog box. Your code should handle this event, as shown in the next section.

Consider a scenario in which the UI of the dialog should correlate with the currently active worksheet and that worksheet's position relative to the other worksheets. In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box. In this case the current worksheet is named "My Sheet" and it is the 2nd sheet in the workbook. The data is encapsulated in an object which is stringified so that it can be passed to `messageChild`.

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## Handle DialogParentMessageReceived in the dialog box

In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method. This is typically done in the [Office.onReady or Office.initialize
methods](initialize-add-in.md). The following is an example:

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Then, define the `onMessageFromParent` handler. The following code continues the example from the preceding section. Note that Office passes an argument to the handler and that the `message` property of argument object contains the string from the host page. In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

It is a best practice to verify that your handler is properly registered. You can do this by passing a callback to the `addHandlerAsync` method that runs when the attempt to register the handler completes. Use the handler to log or show an error if the handler was not successfully registered. The following is an example. Note that `reportError` is a function, not defined here, that logs or displays the error.

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## Conditional messaging

Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages. You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](dialog-api-in-office-add-ins.md#conditional-messaging).
