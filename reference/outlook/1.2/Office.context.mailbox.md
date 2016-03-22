 

# mailbox

## [Office](Office.md)[.context](Office.context.md). mailbox

Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|

### Namespaces

[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.

[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.

[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.

### Members

#### ewsUrl :String

Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.

The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx).

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|

### Methods

####  convertToLocalClientTime(timeValue) → {[LocalClientTime](simple-types.md#localclienttime)}

Gets a dictionary containing time information in local client time.

The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.

If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`timeValue`| Date|A Date object|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Returns:

<dl class="param-type">

<dt>Type</dt>

<dd>[LocalClientTime](simple-types.md#localclienttime)</dd>

</dl>

####  convertToUtcClientTime(input) → {Date}

Gets a Date object from a dictionary containing time information.

The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`input`| [LocalClientTime](simple-types.md#localclienttime)|The local time value to convert.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Returns:

A Date object with the time expressed in UTC.

<dl class="param-type">

<dt>Type</dt>

<dd>Date</dd>

</dl>

####  displayAppointmentForm(itemId)

Displays an existing calendar appointment.

The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.

In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.

In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.

If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`itemId`| String|The Exchange Web Services (EWS) identifier for an existing calendar appointment.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  displayMessageForm(itemId)

Displays an existing message.

The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.

In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.

If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.

Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`itemId`| String|The Exchange Web Services (EWS) identifier for an existing message.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### displayNewAppointmentForm(parameters)

Displays a form for creating a new calendar appointment.

The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.

In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.

In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.

If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`parameters`| Object|A dictionary of parameters describing the new appointment.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Description</th></tr></thead><tbody><tr><td><code>requiredAttendees</code></td><td>Array.&lt;String&gt; &#124; Array.&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>An array of strings containing the email addresses or an array containing an <code>EmailAddressDetails</code> object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</td></tr><tr><td><code>optionalAttendees</code></td><td>Array.&lt;String&gt; &#124; Array.&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>An array of strings containing the email addresses or an array containing an EmailAddressDetails object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</td></tr><tr><td><code>start</code></td><td>Date</td><td>A Date object specifying the start date and time of the appointment.</td></tr><tr><td><code>end</code></td><td>Date</td><td>A Date object specifying the end date and time of the appointment.</td></tr><tr><td><code>location</code></td><td>String</td><td>A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</td></tr><tr><td><code>resources</code></td><td>Array.&lt;String&gt;</td><td>An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</td></tr><tr><td><code>subject</code></td><td>String</td><td>A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</td></tr><tr><td><code>body</code></td><td>String</td><td>The body of the appointment message. The body content is limited to a maximum size of 32 KB.</td></tr></tbody></table>|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'], 
    optionalAttendees: ['sam@contoso.com'], 
    start: start, 
    end: end, 
    location: 'Home', 
    resources: ['projector@contoso.com'], 
    subject: 'meeting', 
    body: 'Hello World!'
  });
```

#### getCallbackTokenAsync(callback, [userContext])

Gets a string that contains a token used to get an attachment or item from an Exchange Server.

The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.

You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://msdn.microsoft.com/en-us/library/office/aa494316.aspx) or [GetItem](https://msdn.microsoft.com/en-us/library/office/aa565934.aspx) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx).

Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. 

The token is provided as a string in the `asyncResult.value` property.|
|`userContext`| Object| &lt;optional&gt;|Any state data that is passed to the asynchronous method.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Read|

##### Example

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  getUserIdentityTokenAsync(callback, [userContext])

Gets a token identifying the user and the Office Add-in.

The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://msdn.microsoft.com/EN-US/library/office/fp179828.aspx).

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. 

The token is provided as a string in the `asyncResult.value` property.|
|`userContext`| Object| &lt;optional&gt;|Any state data that is passed to the asynchronous method.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  makeEwsRequestAsync(data, callback, [userContext])

Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.

The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.

You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.

The XML request must specify UTF-8 encoding.

```
<?xml version="1.0" encoding="utf-8"?>
```

Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx).

**NOTE**: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.

#### Version differences

When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`data`| String||The EWS request.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. 

The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.|
|`userContext`| Object| &lt;optional&gt;|Any state data that is passed to the asynchronous method.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadWriteMailbox|
|Applicable Outlook mode| Compose or read|

##### Example

The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var request = 
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```
