---
title: Office.context.mailbox.userProfile - requirement set 1.6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
---

# userProfile

### [Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile

Provides information about the user in an Outlook add-in.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Minimum permission level](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

## Properties

| Property | Minimum<br>permission level | Modes | Return type | Minimum<br>requirement set |
|---|---|---|---|:---:|
| [accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype) | ReadItem | Compose<br>Read | String | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayName](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#displayname) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [emailAddress](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#emailaddress) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#timezone) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
