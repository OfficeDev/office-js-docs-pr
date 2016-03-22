 

# Simple Types

####  AsyncResult

An object which encapsulates the result of an asynchronous request, including status and error information if the request failed.

##### Properties:

|Name| Type| Description|
|---|---|---|
|`asyncContext`| Object|Gets the object passed to the optional `asyncContext` parameter of the invoked method in the same state as it was passed in.|
|`error`| Error|Gets an Error object that provides a description of the error, if any error occurred.|
|`status`| [Office.AsyncResultStatus](Office.md#.AsyncResultStatus)|Gets the status of the asynchronous operation.|
|`value`| Object|Gets the payload or content of this asynchronous operation, if any.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|Applicable Outlook mode| Compose or read|
#### AttachmentDetails

Represents an attachment on an item from the server. Read mode only.

An array of `AttachmentDetail` objects is returned as the `attachments` property of an `Appointment` or `Message` object.

##### Properties:

|Name| Type| Description|
|---|---|---|
|`attachmentType`| [Office.MailboxEnums.AttachmentType](Office.MailboxEnums.md#attachmenttype-string)|Gets a value that indicates the type of an attachment.|
|`contentType`| String|Gets the MIME content type of the attachment.|
|`id`| String|Gets the Exchange attachment ID of the attachment.|
|`isInline`| Boolean|Gets a value that indicates whether the attachment should be displayed in the body of the item.|
|`name`| String|Gets the name of the attachment.|
|`size`| Number|Gets the size of the attachment in bytes.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|
#### Contact

Represents a contact stored on the server. Read mode only.

The list of contacts associated with an email message or appointment is returned in the `contacts` property of the [`Entities`](simple-types.md#entities) object that is returned by the `getEntities` or `getEntitiesByType` method of the active item.

##### Properties:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;nullable&gt;|An array of strings containing the mailing and street addresses associated with the contact.|
|`businessName`| String| &lt;nullable&gt;|A string containing the name of the business associated with the contact.|
|`emailAddresses`| Array.&lt;String&gt;| &lt;nullable&gt;|An array of strings containing the SMTP email addresses associated with the contact.|
|`personName`| String| &lt;nullable&gt;|A string containing the name of the person associated with the contact.|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;nullable&gt;|An array containing a `PhoneNumber` object for each phone number associated with the contact.|
|`urls`| Array.&lt;String&gt;| &lt;nullable&gt;|An array of strings containing the Internet URLs associated with the contact.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Read|
####  EmailAddressDetails

Provides the email properties of the sender or specified recipients of an email message or appointment.

##### Type:

*   Object

##### Properties:

|Name| Type| Description|
|---|---|---|
|`appointmentResponse`| [Office.MailboxEnums.ResponseType](Office.MailboxEnums.md#responsetype-string)|Gets the response that an attendee returned for an appointment. This property applies to only an attendee of an appointment, as represented by the [`optionalAttendees`](Office.context.mailbox.item.md#optionalAttendees) or [`requiredAttendees`](Office.context.mailbox.item.md#requiredAttendees) property. This property returns `undefined` in other scenarios.|
|`displayName`| String|Gets the display name associated with an email address.|
|`emailAddress`| String|Gets the SMTP email address.|
|`recipientType`| [Office.MailboxEnums.RecipientType](Office.MailboxEnums.md#recipienttype-string)|Gets the email address type of a recipient.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|
#### EmailUser

Represents an email account on an Exchange Server.

##### Properties:

|Name| Type| Description|
|---|---|---|
|`displayName`| String|Gets the display name associated with an email address.|
|`emailAddress`| String|Gets the SMTP email address.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|
#### Entities

Represents a collection of entities found in an email message or appointment. Read mode only.

The `Entities` object is a container for the entity arrays returned by the `getEntities` and `getEntitiesByType` methods when the item (either an email message or an appointment) contains one or more entities that have been found by the server. You can use these entities in your code to provide additional context information to the viewer, such as a map to an address found in the item, or to open a dialer for a phone number found in the item.

If no entities of the type specified in the property are present in the item, the property associated with that entity is `null`. For example, if a message contains a street address and a phone number, the `addresses` property and `phoneNumbers` property would contain information, and the other properties would be `null`.

To be recognized as an address, the string must contain a United States postal address that has at least a subset of the elements of a street number, street name, city, state, and zip code.

To be recognized as a phone number, the string must contain a North American phone number format.

Entity recognition relies on natural language recognition that is based on machine learning of large amounts of data. The recognition of an entity is non-deterministic and success sometimes relies on the particular context in the item.

When the property arrays are returned by the `getEntitiesByType` method, only the property for the specified entity contains data; all other properties are `null`.

##### Properties:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Gets the physical addresses (street or mailing addresses) found in an email message or appointment.|
|`contacts`| Array.&lt;[Contact](simple-types.md#contact)&gt;| &lt;nullable&gt;|Gets the contacts found in an email address or appointment.|
|`emailAddresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Gets the email addresses found in an email message or appointment.|
|`meetingSuggestions`| Array.&lt;[MeetingSuggestion](simple-types.md#meetingsuggestion)&gt;| &lt;nullable&gt;|Gets the meeting suggestions found in an email message.|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;nullable&gt;|Gets the phone numbers found in an email message or appointment.|
|`taskSuggestions`| Array.&lt;[TaskSuggestion](simple-types.md#tasksuggestion)&gt;| &lt;nullable&gt;|Gets the task suggestions found in an email message or appointment.|
|`urls`| Array.&lt;String&gt;| &lt;nullable&gt;|Gets the Internet URLs present in an email message or appointment.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|
#### LocalClientTime

Represents a date and time in the local client's time zone. Read mode only.

##### Properties:

|Name| Type| Description|
|---|---|---|
|`month`| Number|Integer value representing the month, beginning with 0 for January to 11 for December.|
|`date`| Number|Integer value representing the day of the month.|
|`year`| Number|Integer value repesenting the year.|
|`hours`| Number|Integer value representing the hour on a 24-hour clock.|
|`minutes`| Number|Integer value representing the minutes.|
|`seconds`| Number|Integer value representing the seconds.|
|`milliseconds`| Number|Integer value representing the milliseconds.|
|`timezoneOffset`| Number|Integer value representing the number of minutes difference between the local time zone and UTC.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|
#### MeetingSuggestion

Represents a suggested meeting found in an item. Read mode only.

The list of meetings suggested in an email message is returned in the `meetingSuggestions` property of the [`Entities`](simple-types.md#entities) object that is returned when the [`getEntities`](Office.context.mailbox.item.md#getEntities) or [`getEntitiesByType`](Office.context.mailbox.item.md#getEntitiesByType) method is called on the active item.

The `start` and `end` values are string representations of a Date object that contains the date and time at which the suggested meeting is to begin and end. The values are in the default time zone specified for the current user.

##### Properties:

|Name| Type| Description|
|---|---|---|
|`attendees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|Gets the attendees for a suggested meeting.|
|`end`| String|Gets the date and time that a suggested meeting is to end.|
|`location`| String|Gets the location of a suggested meeting.|
|`meetingString`| String|Gets a string that was identified as a meeting suggestion.|
|`start`| String|Gets the date and time that a suggested meeting is to begin.|
|`subject`| String|Gets the subject of a suggested meeting.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|
####  NotificationMessageDetails

An array of `NotificationMessageDetails` objects are returned by the [`NotificationMessages.getAllAsync`](NotificationMessages.md#getAllAsync) method.

##### Type:

*   Object

##### Properties:

|Name| Type| Description|
|---|---|---|
|`key`| String|The identifier for the notification message.|
|`type`| [Office.MailboxEnums.ItemNotificationMessageType](Office.MailboxEnums.md#.ItemNotificationMessageType)|The type of notification message.|
|`icon`| String|The resource identifier of the icon used for the message. Only applicable when `type` is `InformationalMessage`.|
|`message`| String|This is the text of the message. Maximum length is 150 characters.|
|`persistent`| Boolean|If `true`, the message remains until removed by this add-in or dismissed by the user. If `false`, it is removed when the user navigates to a different item. Only applicable when `type` is `InformationalMessage`.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|
#### PhoneNumber

Represents a phone number identified in an item. Read mode only.

An array of `PhoneNumber` objects containing the phone numbers found in an email message is returned in the `phoneNumbers` property of the [`Entities`](simple-types.md#entities) object that is returned when you call the [`getEntities`](Office.context.mailbox.item.md#getEntities) method on the selected item.

##### Type:

*   Object

##### Properties:

|Name| Type| Description|
|---|---|---|
|`originalPhoneString`| String|Gets the text that was identified in an item as a phone number.|
|`phoneString`| String|Gets a string containing a phone number. This string contains only the digits of the telephone number and excludes characters like parentheses and hyphens, if they exist in the original item.|
|`type`| String|Gets a string that identifies the type of phone number: `Home`, `Work`, `Mobile`, `Unspecified`.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|
#### TaskSuggestion

Represents a suggested task identified in an item. Read mode only.

The list of tasks suggested in an email message is returned in the `taskSuggestions` property of the [`Entities`][`Entities`](simple-types.md#entities) object that is returned when the [`getEntities`](Office.context.mailbox.item.md#getEntities) or [`getEntitiesByType`](Office.context.mailbox.item.md#getEntitiesByType) method is called on the active item.

##### Properties:

|Name| Type| Description|
|---|---|---|
|`assignees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|Gets the users that should be assigned a suggested task.|
|`taskString`| String|Gets the text of an item that was identified as a task suggestion.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|
