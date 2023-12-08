---
title: Outlook JavaScript APIs supported in Outlook on mobile devices
description: Learn which Outlook JavaScript APIs are supported in Outlook on mobile devices.
ms.date: 12/12/2023
ms.localizationpriority: medium
---

# Outlook JavaScript APIs supported in Outlook on mobile devices

Outlook on Android and on iOS support up to [Mailbox requirement set 1.5](/javascript/api/outlook?view=outlook-js-1.5&preserve-view=true). To further extend the capabilities of an Outlook mobile add-in, certain APIs from later requirement sets, previously available only to Outlook desktop and web clients, are now enabled for mobile support. This article outlines the APIs supported in Outlook mobile and any implementation exceptions.

## Supported APIs

The following table lists a subset of APIs from requirement sets beyond 1.5 that can now be implemented in Outlook mobile add-ins. Even if the minimum requirement set specified in the manifest of your add-in is greater than 1.5, as long as the API used from the later requirement set is supported, the add-in will appear and activate in Outlook on Android or on iOS. For more information on how to specify the minimum requirement set in your add-in, see [Outlook JavaScript API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

|API|Minimum requirement set|Supported Outlook modes|Supported Outlook on mobile clients|
|---|---|---|---|
|[Office.context.mailbox.item.addFileAttachmentFromBase64Async](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-addfileattachmentfrombase64async-member(1))|Mailbox 1.8|Message Compose|- Android<br>- iOS (rollout in progress)|
|[Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body#outlook-office-body-setsignatureasync-member(1))|Mailbox 1.10|Message Compose|Android|
|[Office.context.mailbox.item.disableClientSignatureAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-disableclientsignatureasync-member(1))|Mailbox 1.10|Message Compose|Android|
|[Office.context.mailbox.item.from.getAsync](/javascript/api/outlook/office.from#outlook-office-from-getasync-member(1))|Mailbox 1.7|Message Compose|- Android<br>- iOS (rollout in progress)|
|[Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getcomposetypeasync-member(1))|Mailbox 1.10|Message Compose|Android|

## Unsupported APIs

Although Outlook mobile supports up to requirement set 1.5, there are some APIs from these earlier requirement sets that aren't supported. The following table lists these APIs and also notes features that aren't supported in certain Outlook modes.

|API|Minimum requirement set|Unsupported Outlook modes|
|---|---|---|
|[Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#office-office-context-officetheme-member)|Mailbox preview|- Message Read<br>- Message Compose<br>- Appointment Attendee<br>- Appointment Organizer|
|[Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-ewsurl-member)|Mailbox 1.1|- Message Read<br>- Message Compose<br>- Appointment Attendee<br>- Appointment Organizer|
|[Office.context.mailbox.convertToEwsId](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-converttoewsid-member(1))|Mailbox 1.3|- Message Read<br>- Message Compose<br>- Appointment Attendee<br>- Appointment Organizer|
|[Office.context.mailbox.convertToRestId](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-converttorestid-member(1))|Mailbox 1.3|- Message Read<br>- Message Compose<br>- Appointment Attendee<br>- Appointment Organizer|
|[Office.context.mailbox.displayAppointmentForm](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-displayappointmentform-member(1))|Mailbox 1.1|- Message Read<br>- Message Compose<br>- Appointment Attendee<br>- Appointment Organizer|
|[Office.context.mailbox.displayMessageForm](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-displaymessageform-member(1))|Mailbox 1.1|- Message Read<br>- Message Compose<br>- Appointment Attendee<br>- Appointment Organizer|
|[Office.context.mailbox.displayNewAppointmentForm](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-displaynewappointmentform-member(1))|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.getCallbackTokenAsync(options, callback)](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getcallbacktokenasync-member(1))|Mailbox 1.5|- Message Compose<br>- Appointment Organizer|
|[Office.context.mailbox.getCallbackTokenAsync(callback, userContext)](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getcallbacktokenasync-member(2))|Mailbox 1.1 (Read mode support)<br><br>Mailbox 1.3 (Compose mode support)|- Message Read<br>- Message Compose<br>- Appointment Attendee<br>- Appointment Organizer|
|[Office.context.mailbox.makeEwsRequestAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-makeewsrequestasync-member(1))|Mailbox 1.1|- Message Read<br>- Message Compose<br>- Appointment Attendee<br>- Appointment Organizer|
|[Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.item.displayReplyAllForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.item.displayReplyForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.item.getRegexMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.item.getRegexMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)|Mailbox 1.1|- Message Read<br>- Appointment Attendee|
|[Office.context.mailbox.item.bcc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))<br><br>[Office.context.mailbox.item.cc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))<br><br>[Office.context.mailbox.item.to.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))|Mailbox 1.1|- Message Compose|
|[Office.context.mailbox.item.bcc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))<br><br>[Office.context.mailbox.item.cc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))<br><br>[Office.context.mailbox.item.to.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))|Mailbox 1.1|- Message Compose|
|[Office.context.mailbox.item.body.getTypeAsync](/javascript/api/outlook/office.body#outlook-office-body-gettypeasync-member(1))|Mailbox 1.1|- Message Compose|
|[Office.context.mailbox.item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))|Mailbox 1.1|- Message Compose|
|[Office.context.mailbox.item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))|Mailbox 1.1|- Message Compose|
|[Office.context.mailbox.item.subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))|Mailbox 1.1|- Message Compose|

## See also

- [Outlook JavaScript API requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Add-ins for Outlook on mobile devices](outlook-mobile-addins.md)
- [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md)
- [Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
