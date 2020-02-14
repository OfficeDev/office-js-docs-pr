---
title: Office.context.mailbox.diagnostics - preview requirement set
description: ''
ms.date: 12/16/2019
localization_priority: Normal
---

# diagnostics

### [Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics

Provides diagnostic information to an Outlook add-in.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|[Applicable Outlook mode](../../../outlook/outlook-add-ins-overview.md#extension-points)| Compose or Read|

## Properties

| Property | Minimum<br>permission level | Modes | Return type | Minimum<br>requirement set |
|---|---|---|---|:---:|
| [hostName](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview#hostname) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [hostVersion](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview#hostversion) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [OWAView](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview#owaview) | ReadItem | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
