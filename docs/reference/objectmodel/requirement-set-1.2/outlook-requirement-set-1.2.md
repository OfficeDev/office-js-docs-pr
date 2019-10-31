---
title: Outlook add-in API requirement set 1.2
description: ''
ms.date: 10/30/2019
localization_priority: Normal
---

# Outlook add-in API requirement set 1.2

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set. 

## What's new in 1.2?

Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.

### Change log

- Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.
- Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.
- Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Added `attachments` property to the `formData` parameter.
- Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Added `attachments` property to the `formData` parameter.

## See also

- [Outlook add-ins](/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/outlook/add-ins/quick-start)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
