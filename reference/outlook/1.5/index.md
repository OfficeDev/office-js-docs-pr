# Outlook add-in API requirement set 1.5

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

## What's new in 1.5?

Requirement set 1.5 includes all of the features of [Requirement set 1.4](../1.4/index.md). It added the following features.

- Added support for [pinnable taskpanes](../../../docs/outlook/manifests/pinnable-taskpane.md).
- Added support for calling [REST APIs](../../../docs/outlook/use-rest-api.md).
- Added ability to mark an attachment as inline.
- Added ability to close a taskpane or dialog.

### Change log

- Added [Office.context.mailbox.addHandlerAsync](https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox?product=outlook&version=v1.5#addhandlerasynceventtype-handler-options-callback): Adds an event handler for a supported event.
- Added [Office.EventType](https://dev.office.com/reference/add-ins/outlook/1.5/Office?product=outlook&version=v1.5#eventtype-string): Specifies the event associated with an event handler.
- Added [Office.context.mailbox.restUrl](https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox?product=outlook&version=v1.5#resturl-string): Gets the URL of the REST endpoint for this email account.
- Modified [Office.context.mailbox.getCallbackTokenAsync](https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox?product=outlook&version=v1.5#getcallbacktokenasyncoptions-callback): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.
- Added [Office.context.ui.closeContainer](https://dev.office.com/reference/add-ins/shared/officeui.closecontainer?product=outlook&version=v1.5): 
- Modified [Office.context.mailbox.item.addFileAttachmentAsync](https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox.item?product=outlook&version=v1.5#addfileattachmentasyncuri-attachmentname-options-callback): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyAllForm](https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox.item?product=outlook&version=v1.5#displayreplyallformformdata): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyForm](https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox.item?product=outlook&version=v1.5#displayreplyformformdata): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.

## Additional resources

- [Outlook add-ins](../../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://developer.microsoft.com/en-us/outlook/code-samples)
- [Get started](https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial)
