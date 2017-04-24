# Outlook add-in API requirement set 1.5

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

## What's new in 1.5?

Requirement set 1.5 includes all of the features of [Requirement set 1.4](../1.4/index.md). It added the following features.

- Added support for [pinnable taskpanes](../../../docs/outlook/manifests/pinnable-taskpane.md).
- Added support for calling [REST APIs](../../../docs/outlook/use-rest-api.md).
- Added ability to mark an attachment as inline.
- Added ability to close a taskpane or dialog.

### Change log

- Added [Office.context.mailbox.addHandlerAsync](Office.context.mailbox.md#addHandlerAsync): Adds an event handler for a supported event.
- Added [Office.EventType](Office.md#EventType): Specifies the event associated with an event handler.
- Added [Office.context.mailbox.restUrl](Office.context.mailbox.md#restUrl): Gets the URL of the REST endpoint for this email account.
- Modified [Office.context.mailbox.getCallbackTokenAsync](Office.context.mailbox.md#getCallbackTokenAsync): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.
- Added [Office.context.ui.closeContainer](Office.context.ui.md#closeContainer): 
- Modified [Office.context.mailbox.item.addFileAttachmentAsync](Office.context.mailbox.item.md#addFileAttachmentAsync): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyAllForm](Office.context.mailbox.item.md#displayReplyAllForm): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyForm](Office.context.mailbox.item.md#displayReplyForm): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.

## Additional resources

- [Outlook add-ins](../../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Get started](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
