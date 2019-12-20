---
title: Outlook add-in API requirement set 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
---

# Outlook add-in API requirement set 1.5

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.

## What's new in 1.5?

Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.

- Added support for [pinnable task panes](/outlook/add-ins/pinnable-taskpane).
- Added support for calling [REST APIs](/outlook/add-ins/use-rest-api).
- Added ability to mark an attachment as inline.
- Added ability to close a task pane or dialog.

### Change log

- Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.
- Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.
- Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.
- Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.
- Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.
- Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.

## See also

- [Outlook add-ins](/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/outlook/add-ins/quick-start)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
