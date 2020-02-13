---
title: Office.context.mailbox.userProfile - requirement set 1.2
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
|[Minimum permission level](../../../outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|[Applicable Outlook mode](../../../outlook/outlook-add-ins-overview.md#extension-points)| Compose or Read|

## Properties

| Property | Minimum<br>permission level | Modes | Return type | Minimum<br>requirement set |
|---|---|---|---|:---:|
| [displayName](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#displayname) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [emailAddress](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#emailaddress) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2#timezone) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
