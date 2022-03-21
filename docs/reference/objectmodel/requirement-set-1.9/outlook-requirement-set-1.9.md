---
title: Outlook add-in API requirement set 1.9
description: 'Requirement set 1.9 for Outlook add-in API.'
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.9

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.9?

Requirement set 1.9 includes all of the features of [requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). It added the following features.

- Added new APIs for append-on-send, custom properties, and display form features.
- Added support for `Dialog.messageChild`.

### Change log

- Added [CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#outlook-office-customproperties-getall-member(1)): Adds a new function to the `CustomProperties` object that gets all custom properties.
- Added [Dialog.messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box): Adds a new method that delivers a message from the host page, such as a task pane or a UI-less function file, to a dialog that was opened from the page.
- Added [ExtendedPermissions manifest element](/javascript/api/extendedpermissions): Adds a child element to the [VersionOverrides](/javascript/api/versionoverrides) manifest element. For an add-in to support the [append-on-send feature](../../../outlook/append-on-send.md), the `AppendOnSend` extended permission must be included in the collection of extended permissions.
- Added [Office.context.mailbox.displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displayappointmentformasync-member(1)): Adds a new function to the `Mailbox` object that displays an existing appointment. This is the async version of the `displayAppointmentForm` method.
- Added [Office.context.mailbox.displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaymessageformasync-member(1)): Adds a new function to the `Mailbox` object that displays an existing message. This is the async version of the `displayMessageForm` method.
- Added [Office.context.mailbox.displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaynewappointmentformasync-member(1)): Adds a new function to the `Mailbox` object that displays a new appointment form. This is the async version of the `displayNewAppointmentForm` method.
- Added [Office.context.mailbox.displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaynewmessageformasync-member(1)): Adds a new function to the `Mailbox` object that displays a new message form. This is the async version of the `displayNewMessageForm` method.
- Added [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#outlook-office-body-appendonsendasync-member(1)): Adds a new function to the `Body` object that appends data to the end of the item body in Compose mode.
- Added [Office.context.mailbox.item.displayReplyAllFormAsync](office.context.mailbox.item.md#methods): Adds a new function to the `Item` object that displays the "Reply all" form in Read mode. This is the async version of the `displayReplyAllForm` method.
- Added [Office.context.mailbox.item.displayReplyFormAsync](office.context.mailbox.item.md#methods): Adds a new function to the `Item` object that displays the "Reply" form in Read mode. This is the async version of the `displayReplyForm` method.

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
