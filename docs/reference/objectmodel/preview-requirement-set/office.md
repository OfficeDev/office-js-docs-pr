---
title: Office namespace - preview requirement set
description: ''
ms.date: 08/13/2019
localization_priority: Normal
---

# Office

The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

##### Members and methods

| Member | Type |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Member |
| [CoercionType](#coerciontype-string) | Member |
| [EventType](#eventtype-string) | Member |
| [SourceProperty](#sourceproperty-string) | Member |

### Namespaces

[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.

### Members

#### AsyncResultStatus: String

Specifies the result of an asynchronous call.

##### Type

*   String

##### Properties:

|Name| Type| Description|
|---|---|---|
|`Succeeded`| String|The call succeeded.|
|`Failed`| String|The call failed.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

<br>
---
---

#### CoercionType: String

Specifies how to coerce data returned or set by the invoked method.

##### Type

*   String

##### Properties:

|Name| Type| Description|
|---|---|---|
|`Html`| String|Requests the data be returned in HTML format.|
|`Text`| String|Requests the data be returned in text format.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|

<br>
---
---

#### EventType: String

Specifies the event associated with an event handler.

##### Type

*   String

##### Properties:

| Name | Type | Description | Minimum requirement set |
|---|---|---|---|
|`AppointmentTimeChanged`| String | The date or time of the selected appointment or series has changed. | 1.7 |
|`AttachmentsChanged`| String | An attachment has been added to or removed from the item. | Preview |
|`EnhancedLocationsChanged`| String | The location of the selected appointment has changed. | Preview |
|`ItemChanged`| String | A different Outlook item is selected for viewing while the task pane is pinned. | 1.5 |
|`OfficeThemeChanged`| String | The Office theme on the mailbox has changed. | Preview |
|`RecipientsChanged`| String | The recipient list of the selected item or appointment location has changed. | 1.7 |
|`RecurrenceChanged`| String | The recurrence pattern of the selected series has changed. | 1.7 |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read |

<br>
---
---

#### SourceProperty: String

Specifies the source of the data returned by the invoked method.

##### Type

*   String

##### Properties:

|Name| Type| Description|
|---|---|---|
|`Body`| String|The source of the data is from the body of a message.|
|`Subject`| String|The source of the data is from the subject of a message.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Applicable Outlook mode](/outlook/add-ins/#extension-points)| Compose or Read|
