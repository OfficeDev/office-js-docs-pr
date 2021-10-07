---
title: Office namespace - requirement set 1.1
description: 'Office namespace members available for Outlook add-ins using Mailbox API requirement set 1.1.'
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# Office (Mailbox requirement set 1.1)

The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office?view=outlook-js-1.1&preserve-view=true).

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](../../../outlook/outlook-add-ins-overview.md#extension-points)| Compose or Read|

## Properties

| Property | Modes | Return type | Minimum<br>requirement set |
|---|---|---|:---:|
| [context](office.context.md) | Compose<br>Read | [Context](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Enumerations

| Enumeration | Modes | Return type | Minimum<br>requirement set |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [SourceProperty](#sourceproperty-string) | Compose<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## Namespaces

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.

## Enumeration details

#### AsyncResultStatus: String

Specifies the result of an asynchronous call.

##### Type

*   String

##### Properties

|Name| Type| Description|
|---|---|---|
|`Succeeded`| String|The call succeeded.|
|`Failed`| String|The call failed.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](../../../outlook/outlook-add-ins-overview.md#extension-points)| Compose or Read|

<br>

---
---

#### CoercionType: String

Specifies how to coerce data returned or set by the invoked method.

##### Type

*   String

##### Properties

|Name| Type| Description|
|---|---|---|
|`Html`| String|Requests the data be returned in HTML format.|
|`Text`| String|Requests the data be returned in text format.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](../../../outlook/outlook-add-ins-overview.md#extension-points)| Compose or Read|

<br>

---
---

#### SourceProperty: String

Specifies the source of the data returned by the invoked method.

##### Type

*   String

##### Properties

|Name| Type| Description|
|---|---|---|
|`Body`| String|The source of the data is from the body of a message.|
|`Subject`| String|The source of the data is from the subject of a message.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Applicable Outlook mode](../../../outlook/outlook-add-ins-overview.md#extension-points)| Compose or Read|