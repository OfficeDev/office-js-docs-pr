

# item

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).item

The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype) property.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Applicable Outlook mode| Compose or read|

### Example

The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### Members

#### attachments :Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

Gets an array of attachments for the item. Read mode only.

##### Type:

*   Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

The following code builds an HTML string with details of all attachments on the current item.

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  bcc :[Recipients](Recipients.md)

Gets or sets the recipients on the Bcc (blind carbon copy) line of a message. Compose mode only.

##### Type:

*   [Recipients](Recipients.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose|

##### Example

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  body :[Body](Body.md)

Gets an object that provides methods for manipulating the body of an item.

##### Type:

*   [Body](Body.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|
####  cc :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Gets or sets the Cc (carbon copy) recipients of a message.

##### Read mode

The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.

##### Compose mode

The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.

##### Type:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  (nullable) conversationId :String

Gets an identifier for the email conversation that contains a particular message.

You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.

You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|
#### dateTimeCreated :Date

Gets the date and time that an item was created. Read mode only.

##### Type:

*   Date

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### dateTimeModified :Date

Gets the date and time that an item was last modified. Read mode only.

> **Note:** This member is not supported in Outlook for iOS or Outlook for Android.

##### Type:

*   Date

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  end :Date|[Time](Time.md)

Gets or sets the date and time that the appointment is to end.

The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.

##### Read mode

The `end` property returns a `Date` object.

##### Compose mode

The `end` property returns a `Time` object.

When you use the [`Time.setAsync`](Time.md#setasyncdatetime-options-callback) method to set the end time, you should use the [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.

##### Type:

*   Date | [Time](Time.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

The following example sets the end time of an appointment in compose mode by using the [`setAsync`](Time.md#setasyncdatetime-options-callback) method of the `Time` object.

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
	 asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### from :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Gets the email address of the sender of a message. Read mode only.

The `from` and [`sender`](Office.context.mailbox.item.md#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.

> **Note:** The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.

##### Type:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|
#### internetMessageId :String

Gets the Internet message identifier for an email message. Read mode only.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### itemClass :String

Gets the Exchange Web Services item class of the selected item. Read mode only.

The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.

| Type | Description | item class |
| --- | --- | --- |
| Appointment items | These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Message items | These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### (nullable) itemId :String

Gets the Exchange Web Services item identifier for the current item. Read mode only.

> **Note:** The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier. The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API. Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](Office.context.mailbox.md). For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).

The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](simple-types.md#asyncresult) parameter in the callback function.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  itemType :[Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

Gets the type of item that an instance represents.

The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.

##### Type:

*   [Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  location :String|[Location](Location.md)

Gets or sets the location of an appointment.

##### Read mode

The `location` property returns a string that contains the location of the appointment.

##### Compose mode

The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.

##### Type:

*   String | [Location](Location.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### normalizedSubject :String

Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.

The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](Office.context.mailbox.item.md#subject-stringsubject) property.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  notificationMessages :[NotificationMessages](NotificationMessages.md)

Gets the notification messages for an item.

##### Type:

*   [NotificationMessages](NotificationMessages.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|
|[Recipients](Recipients.md)|
####  optionalAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>
Gets or sets a list of email addresses for optional attendees.

##### Read mode

The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.

##### Compose mode

The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.

##### Type:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### organizer :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Gets the email address of the meeting organizer for a specified meeting. Read mode only.

##### Type:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  requiredAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Gets or sets a list of email addresses for required attendees.

##### Read mode

The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.

##### Compose mode

The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.

##### Type:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### resources :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Gets the resources required for an appointment. Read mode only.

> **Note:** This member is not supported in Outlook for iOS or Outlook for Android.

##### Type:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|
#### sender :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Gets the email address of the sender of an email message. Read mode only.

The [`from`](Office.context.mailbox.item.md#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.

> **Note:** The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.

##### Type:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  start :Date|[Time](Time.md)

Gets or sets the date and time that the appointment is to begin.

The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.

##### Read mode

The `start` property returns a `Date` object.

##### Compose mode

The `start` property returns a `Time` object.

When you use the [`Time.setAsync`](Time.md#setasyncdatetime-options-callback) method to set the start time, you should use the [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.

##### Type:

*   Date | [Time](Time.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

The following example sets the start time of an appointment in compose mode by using the [`setAsync`](Time.md#setasyncdatetime-options-callback) method of the `Time` object.

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
	 asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  subject :String|[Subject](Subject.md)

Gets or sets the description that appears in the subject field of an item.

The `subject` property gets or sets the entire subject of the item, as sent by the email server.

##### Read mode

The `subject` property returns a string. Use the [`normalizedSubject`](Office.context.mailbox.item.md#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.

```
var subject = Office.context.mailbox.item.subject;
```

##### Compose mode

The `subject` property returns a `Subject` object that provides methods to get and set the subject.

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### Type:

*   String | [Subject](Subject.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|
####  to :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Gets or sets the recipients of an email message.

##### Read mode

The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.

##### Compose mode

The `to` property returns a `Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.

##### Type:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### Methods

####  addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Adds a file to a message or appointment as an attachment.

The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.

You can subsequently use the identifier with the [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.

##### Parameters:
|Name| Type| Attributes| Description|
|---|---|---|---|
|`uri`| String||The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.|
|`attachmentName`| String||The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
| `options.asyncContext` | Object | &lt;optional&gt; | Developers can provide any object they wish to access in the callback method. |
| `options.isInline` | Boolean | &lt;optional&gt; | If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list. |
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>On success, the attachment identifier will be provided in the `asyncResult.value` property.<br/>If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.|

##### Errors

| Error code | Description |
|------------|-------------|
| `AttachmentSizeExceeded` | The attachment is larger than allowed. |
| `FileTypeNotSupported` | The attachment has an extension that is not allowed. |
| `NumberOfAttachmentsExceeded` | The message or appointment has too many attachments. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Examples

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

The following example adds an image file as an inline attachment and references the attachment in the message body.

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Adds an Exchange item, such as a message, as an attachment to the message or appointment.

The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.

You can subsequently use the identifier with the [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.

If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`itemId`| String||The Exchange identifier of the item to attach. The maximum length is 100 characters.|
|`attachmentName`| String||The sujbect of the item to be attached. The maximum length is 255 characters.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>On success, the attachment identifier will be provided in the `asyncResult.value` property.<br/>If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.|

##### Errors

| Error code | Description |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | The message or appointment has too many attachments. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Example

The following example adds an existing Outlook item as an attachment with the name `My Attachment`.

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  close()

Closes the current item that is being composed.

The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.

In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Applicable Outlook mode| Compose|

#### displayReplyAllForm(formData)

Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.

If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.

When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.

##### Parameters:

| Name | Type | Attributes | Description |
|---|---|---|---|
|`formData`| String &#124; Object| |A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.<br/>**OR**<br/>An object that contains body or attachment data and a callback function. The object is defined as follows. |
| `formData.htmlBody` | String | &lt;optional&gt; | A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;optional&gt; | An array of JSON objects that are either file or item attachments. |
| `formData.attachments[].type` | String | | Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment. |
| `formData.attachments[].name` | String | | A string that contains the name of the attachment, up to 255 characters in length.| 
| `formData.attachments[].url` | String | | Only used if `type` is set to `file`. The URI of the location for the file. |
| `formData.attachments[].isInline` | Boolean | | Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list. |
| `formData.attachments[].itemId` | String | | Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters. |
| `callback` | function | &lt;optional&gt; | When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](simple-types.md#asyncresult) object. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Examples

The following code passes a string to the `displayReplyAllForm` function.

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Reply with an empty body.

```
Office.context.mailbox.item.displayReplyAllForm({});
```

Reply with just a body.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Reply with a body and a file attachment.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Reply with a body and an item attachment.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Reply with a body, file attachment, item attachment, and a callback.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### displayReplyForm(formData)

Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.

If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.

When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.

##### Parameters:

| Name | Type | Attributes | Description |
|---|---|---|---|
|`formData`| String &#124; Object| | A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.<br/>**OR**<br/>An object that contains body or attachment data and a callback function. The object is defined as follows. |
| `formData.htmlBody` | String | &lt;optional&gt; | A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;optional&gt; | An array of JSON objects that are either file or item attachments. |
| `formData.attachments[].type` | String | | Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment. |
| `formData.attachments[].name` | String | | A string that contains the name of the attachment, up to 255 characters in length.| 
| `formData.attachments[].url` | String | | Only used if `type` is set to `file`. The URI of the location for the file. |
| `formData.attachments[].isInline` | Boolean | | Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list. |
| `formData.attachments[].itemId` | String | | Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters. |
| `callback` | function | &lt;optional&gt; | When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](simple-types.md#asyncresult) object. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Examples

The following code passes a string to the `displayReplyForm` function.

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Reply with an empty body.

```
Office.context.mailbox.item.displayReplyForm({});
```

Reply with just a body.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Reply with a body and a file attachment.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Reply with a body and an item attachment.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Reply with a body, file attachment, item attachment, and a callback.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### getEntities() → {[Entities](simple-types.md#entities)}

Gets the entities found in the selected item.

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Returns:

Type:
[Entities](simple-types.md#entities)

##### Example

The following example accesses the contacts entities on the current item.

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Gets an array of all the entities of the specified entity type found in the selected item.

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#.EntityType-string)|One of the EntityType enumeration values.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Applicable Outlook mode| Read|

##### Returns:

If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null. If no entities of the specified type are present on the item, the method returns an empty array. Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.

While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.

| Value of `entityType` | Type of objects in returned array | Required Permission Level |
| --- | --- | --- |
| `Address` | String | **Restricted** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restricted** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **Restricted** |

Type: Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

##### Example

The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/office/fp161166.aspx) rule element in the manifest XML file with the specified `FilterName` element value.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The name of the `ItemHasKnownEntity` rule element that defines the filter to match.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Returns:

If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.

Type:
Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

#### getRegExMatches() → {Object}

Returns string values in the selected item that match the regular expressions defined in the manifest XML file.

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.

For example, consider an add-in manifest has the following `Rule` element:

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](Body.md#getasynccoerciontype-options-callback) method to retrieve the entire body.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Returns:

An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.

<dl class="param-type">

<dt>Type</dt>

<dd>Object</dd>

</dl>

##### Example

The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### getRegExMatchesByName(name) → (nullable) {Array.<String>}

Returns string values in the selected item that match the named regular expression defined in the manifest XML file.

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.

If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Returns:

An array that contains the strings that match the regular expression defined in the manifest XML file.

<dl class="param-type">

<dt>Type</dt>

<dd>Array.<String></dd>

</dl>

##### Example

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  getSelectedDataAsync(coercionType, [options], callback) → {String}

Asynchronously returns selected data from the subject or body of a message.

If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.

To access the selected data from the callback method, call `asyncResult.value.data`. To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Returns:

The selected data as a string with format determined by `coercionType`.

<dl class="param-type">

<dt>Type</dt>

<dd>String</dd>

</dl>

##### Example

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### getSelectedEntities() → {[Entities](simple-types.md#entities)}

Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](..\..\..\docs\outlook\contextual-outlook-add-ins.md).

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| Preview |
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Returns:

Type:
[Entities](simple-types.md#entities)

##### Example

The following example accesses the addresses entities in the highlighted match selected by the user.

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### getSelectedRegExMatches() → {Object}

Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](..\..\..\docs\outlook\contextual-outlook-add-ins.md).

> **Note:** This method is not supported in Outlook for iOS or Outlook for Android.

The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.

For example, consider an add-in manifest has the following `Rule` element:

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](Body.md#getasynccoerciontype-options-callback) method to retrieve the entire body.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| Preview |
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Read|

##### Returns:

An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.

##### Example

The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  loadCustomPropertiesAsync(callback, [userContext])

Asynchronously loads custom properties for this add-in on the selected item.

Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.

The custom properties are provided as a [`CustomProperties`](CustomProperties.md) object in the `asyncResult.value` property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.|
|`userContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback function. This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  removeAttachmentAsync(attachmentId, [options], [callback])

Removes an attachment from a message or appointment.

The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`attachmentId`| String||The identifier of the attachment to remove. The maximum length of the string is 100 characters.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.|

##### Errors

| Error code | Description |
|------------|-------------|
| `InvalidAttachmentId` | The attachment identifier does not exist. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Example

The following code removes an attachment with an identifier of '0'.

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  saveAsync([options], callback)

Asynchronously saves an item.

When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.

> **Note:** If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the `itemId` will return an error.

Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.

> **Note:** Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.

On success, the item identifier is provided in the `asyncResult.value` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Examples

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  setSelectedDataAsync(data, [options], callback)

Asynchronously inserts data into the body or subject of a message.

The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`data`| String||The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`options.coercionType`| [Office.CoercionType](Office.md)| &lt;optional&gt;|If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.<br/><br/>If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.<br/><br/>If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.2|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Example

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```