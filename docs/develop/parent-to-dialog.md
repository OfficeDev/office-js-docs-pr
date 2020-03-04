---
title: Passing data and messages to a dialog box from its host page
description: 'Learn how to pass data to a dialog from the host page'
ms.date: 03/11/2020
localization_priority: Normal
---

# Passing data and messages to a dialog box from its host page

Microsoft is now making available for preview new APIs for messaging from the host page (defined [here](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)) to the dialog box.

> [!Important]
>
> - The APIs are in preview. They are available to developers for experimentation; but they should not be used in a production add-in. Until this API is released, use the techniques described in [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box).
> - The APIs described in this article require Office 365 (the subscription version of Office). You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). Please note that when a build graduates to the production semi-annual channel, support for preview features is turned off for that build.
> - The APIs are only supported in Excel during the preview.

## Use messageChild() from the host page

When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned and should be assigned to a variable which typically has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-)
method because the object will be referenced in other methods. The following is an example:

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

This `Dialog` object has a `messageChild` method that will send any string, or stringified data, to the dialog box. This raises a `DialogParentMessageReceived` event in the dialog box. Your code should handle this event. (See next section.)

In the following example, `sheetPropertiesChanged` sends newly changed Excel worksheet properties to the dialog box.

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

In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event. This would typically be done in the [Office.onReady or Office.initialize
methods](initialize-add-in.md). The following is an example:

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Then, of course, define the `onMessageFromParent` handler. The following code continues the example from the preceding section. Note that Office passes an argument to the handler and that the `message` property of argument object contains the string from the host page. In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.

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
    if (asyncResult.status != Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## Conditional messaging

Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages. You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](dialog-api-in-office-add-ins.md#conditional-messaging).
