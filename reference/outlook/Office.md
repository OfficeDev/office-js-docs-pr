 

# Office

The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](../shared/shared-api.md).

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose or read|

### Namespaces

[context](Office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.

[MailboxEnums](Office.MailboxEnums.md): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.

### Members

####  AsyncResultStatus :String

Specifies the result of an asynchronous call.

##### Type:

*   String

##### Properties:

|Name| Type| Description|
|---|---|---|
|`Succeeded`| String|The call succeeded.|
|`Failed`| String|The call failed.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose or read|
####  CoercionType :String

Specifies how to coerce data returned or set by the invoked method.

##### Type:

*   String

##### Properties:

|Name| Type| Description|
|---|---|---|
|`Html`| String|Requests the data be returned in HTML format.|
|`Text`| String|Requests the data be returned in text format.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose or read|
####  SourceProperty :String

Specifies the source of the data returned by the invoked method.

##### Type:

*   String

##### Properties:

|Name| Type| Description|
|---|---|---|
|`Body`| String|The source of the data is from the body of a message.|
|`Subject`| String|The source of the data is from the subject of a message.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose or read|
