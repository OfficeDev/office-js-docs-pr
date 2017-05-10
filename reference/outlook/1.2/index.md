# Outlook add-in API requirement set 1.2

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a [requirement set](../tutorial-api-requirement-sets.md) other than the latest requirement set. 

## What's new in 1.2?

Requirement set 1.2 includes all of the features of [Requirement set 1.1](../1.1/index.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.

### Change log

- Added [Office.context.mailbox.item.setSelectedDataAsync](https://dev.office.com/reference/add-ins/outlook/1.2/Office.context.mailbox.item?product=outlook&version=v1.2#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.
- Modified [Office.context.mailbox.item.displayReplyAllForm](https://dev.office.com/reference/add-ins/outlook/1.2/Office.context.mailbox.item?product=outlook&version=v1.2#displayreplyallformformdata): Added `attachments` property to the `formData` parameter.
- Modified [Office.context.mailbox.item.displayReplyForm](https://dev.office.com/reference/add-ins/outlook/1.2/Office.context.mailbox.item?product=outlook&version=v1.2#displayreplyformformdata): Added `attachments` property to the `formData` parameter.

## Additional resources

- [Outlook add-ins](../../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://developer.microsoft.com/en-us/outlook/code-samples)
- [Get started](https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial)
