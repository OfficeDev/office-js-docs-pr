---
title: Office.context.mailbox.item - requirement set 1.7
description: 'Outlook Mailbox API requirement set 1.7 version of the Item object model.'
ms.date: 07/16/2021
ms.localizationpriority: medium
---

# item (Mailbox requirement set 1.7)

### [Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` is used to access the currently selected message, meeting request, or appointment. You can determine the type of the item by using the `itemType` property.

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Minimum permission level](../../../outlook/understanding-outlook-add-in-permissions.md)|Restricted|
|[Applicable Outlook mode](../../../outlook/outlook-add-ins-overview.md#extension-points)|Appointment Organizer, Appointment Attendee,<br>Message Compose, or Message Read|

> [!IMPORTANT]
> Android and iOS: There are limitations on when add-ins activate and which APIs are available. To learn more, refer to [Add mobile support to an Outlook add-in](../../../outlook/add-mobile-support.md#compose-mode-and-appointments).

## Properties

| Property | Minimum<br>permission level | Details by mode | Return type | Minimum<br>requirement set |
|---|---|---|---|:---:|
| attachments | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-bcc-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-cc-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-cc-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-conversationId-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-conversationId-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-dateTimeCreated-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-dateTimeCreated-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-dateTimeModified-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-dateTimeModified-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-end-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-end-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-end-member)<br>(Meeting Request) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadWriteItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-from-member) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-from-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-internetMessageId-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-itemClass-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-itemClass-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-itemId-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-itemId-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-itemType-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-itemType-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-itemType-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-itemType-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-location-member) | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-location-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-location-member)<br>(Meeting Request) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-normalizedSubject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-normalizedSubject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-notificationMessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-notificationMessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-notificationMessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-notificationMessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-optionalAttendees-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-optionalAttendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-organizer-member) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-organizer-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-recurrence-member) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-recurrence-member) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-recurrence-member)<br>(Meeting Request) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-requiredAttendees-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-requiredAttendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-sender-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-seriesId-member) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-seriesId-member) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-seriesId-member) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-seriesId-member) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-start-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-start-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-start-member)<br>(Meeting Request) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-subject-member) | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-subject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-subject-member) | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-subject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| to | ReadItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-to-member) | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-to-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Methods

| Method | Minimum<br>permission level | Details by mode | Minimum<br>requirement set |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-addFileAttachmentAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-addFileAttachmentAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-addHandlerAsync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-addHandlerAsync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-addHandlerAsync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-addHandlerAsync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-addItemAttachmentAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-addItemAttachmentAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-displayReplyAllForm-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-displayReplyAllForm-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-displayReplyForm-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-displayReplyForm-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities() | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getEntities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getEntities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getEntitiesByType-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getEntitiesByType-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getFilteredEntitiesByName-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getFilteredEntitiesByName-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches() | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getRegExMatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getRegExMatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getRegExMatchesByName-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getRegExMatchesByName-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-getSelectedDataAsync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-getSelectedDataAsync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getSelectedEntities-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getSelectedEntities-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-getSelectedRegExMatches-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-getSelectedRegExMatches-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-loadCustomPropertiesAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-loadCustomPropertiesAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-loadCustomPropertiesAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-loadCustomPropertiesAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-removeAttachmentAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-removeAttachmentAsync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-removeHandlerAsync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentread-removeHandlerAsync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-removeHandlerAsync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#outlook-office-messageread-removeHandlerAsync-member(1)) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-saveAsync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-saveAsync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#outlook-office-appointmentcompose-setSelectedDataAsync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#outlook-office-messagecompose-setSelectedDataAsync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## Events

You can subscribe to and unsubscribe from the following events using `addHandlerAsync` and `removeHandlerAsync` respectively.

> [!IMPORTANT]
> Events are only available with task pane implementation.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.7&preserve-view=true) | Description | Minimum<br>requirement set |
|---|---|:---:|
|`AppointmentTimeChanged`| The date or time of the selected appointment or series has changed. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecipientsChanged`| The recipient list of the selected item or appointment location has changed. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| The recurrence pattern of the selected series has changed. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

## Example

The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```
