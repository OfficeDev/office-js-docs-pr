 

# MailboxEnums

## [Office](Office.md). MailboxEnums

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose|

### Members

#### AttachmentType :String

Specifies an attachment's type. Compose mode only.

AttachmentType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`File`| String|`file`|The attachment is a file.|
|`Item`| String|`item`|The attachment is an Exchange item.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose|
#### EntityType :String

Specifies an entity's type. Compose mode only.

EntityType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`Address`| String|`address`|Specifies that the entity is a postal address.|
|`Contact`| String|`contact`|Specifies that the entity is a contact.|
|`EmailAddress`| String|`emailAddress`|Specifies that the entity is SMTP email address.|
|`MeetingSuggestion`| String|`meetingSuggestion`|Specifies that the entity is a meeting suggestion.|
|`PhoneNumber`| String|`phoneNumber`|Specifies that the entity is US phone number.|
|`TaskSuggestion`| String|`taskSuggestion`|Specifies that the entity is a task suggestion.|
|`URL`| String|`url`|Specifies that the entity is an Internet URL.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose|
#### ItemNotificationMessageType :String

Specifies the notification message type for an appointment or message. Compose mode only.

ItemNotificationMessageType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`ProgressIndicator`| String|`progressIndicator`|The notificationMessage is a progress indicator.|
|`InformationalMessage`| String|`informationalMessage`|The notificationMessage is an informational message.|
|`ErrorMessage`| String|`errorMessage`|The notificationMessage is an error message.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|Applicable Outlook mode| Compose|
#### ItemType :String

Specifies an item's type. Compose mode only.

ItemType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`Message`| String|`message`|An email, meeting request, meeting response, or meeting cancellation.|
|`Appointment`| String|`appointment`|An appointment item.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose|
#### RecipientType :String

Specifies the type of recipient for an appointment. Compose mode only.

RecipientType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`Other`| String|`other`|The recipient is not one of the other recipient types.|
|`DistributionList`| String|`distributionList`|The recipient is a distribution list containing a list of email addresses.|
|`User`| String|`user`|The recipient is an SMTP email address that is on the Exchange server.|
|`ExternalUser`| String|`externalUser`|The recipient is an SMTP email address that is not on the Exchange server.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.1|
|Applicable Outlook mode| Compose|
#### ResponseType :String

Specifies the type of response to a meeting invitation. Compose mode only.

ResponseType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`None`| String|`none`|There has been no response from the attendee.|
|`Organizer`| String|`organizer`|The attendee is the meeting organizer.|
|`Tentative`| String|`tentative`|The meeting request was tentatively accepted by the attendee.|
|`Accepted`| String|`accepted`|The meeting request was accepted by the attendee.|
|`Declined`| String|`declined`|The meeting request was declined by the attendee.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose|
