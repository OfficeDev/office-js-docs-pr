# Outlook add-in API requirement set 1.2

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a **preview** [requirement set](tutorial-api-requirement-sets.html). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest. Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.

The Preview Requirement set includes all of the features of [Requirement set 1.5](../1.5/index.md). 

## Features in preview

The following features are in preview.

- [Office.context.mailbox.addHandlerAsync](Office.context.mailbox.md#addHandlerAsync)
- [Office.EventType](Office.md#EventType)
- [Office.context.mailbox.restUrl](Office.context.mailbox.md#restUrl)
- [Office.context.mailbox.getCallbackTokenAsync](Office.context.mailbox.md#getCallbackTokenAsync) - A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.
- [Office.context.ui.closeContainer](Office.context.ui.md#closeContainer)
- [Office.context.mailbox.item.addFileAttachmentAsync](Office.context.mailbox.item.md#addFileAttachmentAsync) - A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- [Office.context.mailbox.item.displayReplyAllForm](Office.context.mailbox.item.md#displayReplyAllForm) - A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- [Office.context.mailbox.item.displayReplyForm](Office.context.mailbox.item.md#displayReplyForm) - A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- [Event.completed](Event.md#completed) - A new optional parameter `options`, which is a dictionary with one valid value `allowEvent`. This value is used to cancel execution of an event.

## Additional resources

- [Outlook add-ins](../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Get started](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
