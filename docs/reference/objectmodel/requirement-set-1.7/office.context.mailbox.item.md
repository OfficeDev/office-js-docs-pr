---
title: Office.context.mailbox.item - requirement set 1.7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
---

# item

### [Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` is used to access the currently selected message, meeting request, or appointment. You can determine the type of the item by using the `itemType` property.

##### Requirements

|Requirement|Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Minimum permission level](/outlook/add-ins/understanding-outlook-add-in-permissions)|Restricted|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)|Compose or Read|

## Properties

| Property | Minimum<br>permission level | Modes | Return type | Minimum<br>requirement set |
|---|---|---|---|:---:|
| attachments | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#bcc) | [Recipients](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#cc) | [Recipients](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#end)<br>(Meeting Request) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadWriteItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#from) | [From](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#location)<br>(Meeting Request) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#optionalattendees) | [Recipients](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#optionalattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#recurrence)<br>(Meeting Request) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#requiredattendees) | [Recipients](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#requiredattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#start)<br>(Meeting Request) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| to | ReadItem | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#to) | [Recipients](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Methods

| Method | Minimum<br>permission level | Modes | Minimum<br>requirement set |
|---|---|---|:---:|
| addFileAttachmentAsync | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addHandlerAsync | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close | Restricted | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType | Restricted | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches | ReadItem | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync | ReadItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Appointment Attendee](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message Read](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync | ReadWriteItem | [Appointment Organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Message Compose](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## Events

You can subscribe to and unsubscribe from the following events using `addHandlerAsync` and `removeHandlerAsync` respectively.

| Event | Description | Minimum<br>requirement set |
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
