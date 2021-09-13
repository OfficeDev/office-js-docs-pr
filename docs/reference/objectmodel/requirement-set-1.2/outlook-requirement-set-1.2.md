---
title: Outlook add-in API requirement set 1.2
description: 'Features and APIs that were introduced for Outlook Add-ins and the Office JavaScript APIs as part of Mailbox API 1.2.'
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.2

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.2?

Requirement set 1.2 includes all of the features of [requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.

### Change log

- Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.
- Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.
- Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.
- Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
