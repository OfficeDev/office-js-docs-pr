---
title: Office.context.mailbox - preview requirement set
description: ''
ms.date: 12/16/2019
localization_priority: Normal
---

# mailbox

### [Office](office.md)[.context](office.context.md).mailbox

Provides access to the Outlook add-in object model for Microsoft Outlook.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Minimum permission level](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

## Properties

| Property | Minimum<br>permission level | Modes | Return type | Minimum<br>requirement set |
|---|---|---|---|:---:|
| [diagnostics](office.context.mailbox.diagnostics.md) | ReadItem | Compose<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#ewsurl) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | Compose<br>Read | [Item](/javascript/api/outlook/office.item?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#mastercategories) | ReadWriteMailbox | Compose<br>Read | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-preview) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#resturl) | ReadItem | Compose<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](office.context.mailbox.userProfile.md) | ReadItem | Compose<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Methods

| Method | Minimum<br>permission level | Modes | Minimum<br>requirement set |
|---|---|---|:---:|
| [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Compose<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttoewsid-itemid--restversion-) | Restricted | Compose<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttolocalclienttime-timevalue-) | ReadItem | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttorestid-itemid--restversion-) | Restricted | Compose<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttoutcclienttime-input-) | ReadItem | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentform-itemid-) | ReadItem | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageform-itemid-) | ReadItem | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentform-parameters-) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageform-parameters-) | ReadItem | Compose<br>Read | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getcallbacktokenasync-options--callback-) | ReadItem | Compose<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getcallbacktokenasync-callback--usercontext-) | ReadItem | Compose<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Compose<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | ReadItem | Compose<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## Events

You can subscribe to and unsubscribe from the following events using [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) and [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) respectively.

| Event | Description | Minimum<br>requirement set |
|---|---|:---:|
|`ItemChanged`| A different Outlook item is selected for viewing while the task pane is pinned. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
|`OfficeThemeChanged`| The Office theme on the mailbox has changed. | [Preview](../preview-requirement-set/outlook-requirement-set-preview.md) |
