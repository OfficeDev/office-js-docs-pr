
# mailbox

### [Office](Office.md)[.context](Office.context.md). mailbox

Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Members and methods

| Member | Type |
|--------|------|
| [ewsUrl](#ewsurl-string) | Member |
| [restUrl](#resturl-string) | Member |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Method |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | Method |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | Method |
| [convertToRestId](#converttorestiditemid-restversion--string) | Method |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Method |
| [displayAppointmentForm](#displayappointmentformitemid) | Method |
| [displayMessageForm](#displaymessageformitemid) | Method |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | Method |
| [getCallbackTokenAsync](#getcallbacktokenasyncoptions-callback) | Method |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Method |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Method |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Method |

### Namespaces

[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.

[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.

[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.

### Members

#### ewsUrl :String

Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.

> [!NOTE]
> This member is not supported in Outlook for iOS or Outlook for Android.

The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.

In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

#### restUrl :String

Gets the URL of the REST endpoint for this email account.

The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.

Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.

In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.

> [!NOTE]
> Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

### Methods

####  addHandlerAsync(eventType, handler, [options], [callback])

Adds an event handler for a supported event.

Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.

##### Parameters:

| Name | Type | Attributes | Description |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || The event that should invoke the handler. |
| `handler` | Function || The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`. |
| `options` | Object | &lt;optional&gt; | An object literal that contains one or more of the following properties. |
| `options.asyncContext` | Object | &lt;optional&gt; | Developers can provide any object they wish to access in the callback method. |
| `callback` | function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Example

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  convertToEwsId(itemId, restVersion) → {String}

Converts an item ID formatted for REST into EWS format.

> [!NOTE]
> This method is not supported in Outlook for iOS or Outlook for Android.

Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`itemId`| String|An item ID formatted for the Outlook REST APIs|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|A value indicating the version of the Outlook REST API used to retrieve the item ID.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Returns:

Type:
String

##### Example

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}

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
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Returns:

Type:
[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)

####  convertToRestId(itemId, restVersion) → {String}

Converts an item ID formatted for EWS into REST format.

> [!NOTE]
> This method is not supported in Outlook for iOS or Outlook for Android.

Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`itemId`| String|An item ID formatted for Exchange Web Services (EWS)|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|A value indicating the version of the Outlook REST API that the converted ID will be used with.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Returns:

Type:
String

##### Example

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  convertToUtcClientTime(input) → {Date}

Gets a Date object from a dictionary containing time information.

The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)|The local time value to convert.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Returns:

A Date object with the time expressed in UTC.

<dl class="param-type">

<dt>Type</dt>

<dd>Date</dd>

</dl>

####  displayAppointmentForm(itemId)

Displays an existing calendar appointment.

> [!NOTE]
> This method is not supported in Outlook for iOS or Outlook for Android.

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
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Example

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  displayMessageForm(itemId)

Displays an existing message.

> [!NOTE]
> This method is not supported in Outlook for iOS or Outlook for Android.

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
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

##### Example

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### displayNewAppointmentForm(parameters)

Displays a form for creating a new calendar appointment.

> [!NOTE]
> This method is not supported in Outlook for iOS or Outlook for Android.

The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.

In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.

In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.

If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.

##### Parameters:

|Name| Type| Description|
|---|---|---|
| `parameters` | Object | A dictionary of parameters describing the new appointment. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt; | An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt; | An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries. |
| `parameters.start` | Date | A `Date` object specifying the start date and time of the appointment. |
| `parameters.end` | Date | A `Date` object specifying the end date and time of the appointment. |
| `parameters.location` | String | A string containing the location of the appointment. The string is limited to a maximum of 255 characters. |
| `parameters.resources` | Array.&lt;String&gt; | An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries. |
| `parameters.subject` | String | A string containing the subject of the appointment. The string is limited to a maximum of 255 characters. |
| `parameters.body` | String | The body of the appointment. The body content is limited to a maximum size of 32 KB. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Read|

##### Example

```js
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

#### getCallbackTokenAsync([options], callback)

Gets a string that contains a token used to call REST APIs or Exchange Web Services.

The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.

> [!NOTE]
> It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible. 

**REST Tokens**

When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.

The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.

**EWS Tokens**

When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.

The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
| `options` | Object | &lt;optional&gt; | An object literal that contains one or more of the following properties. |
| `options.isRest` | Boolean |  &lt;optional&gt; | Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`. |
| `options.asyncContext` | Object |  &lt;optional&gt; | Any state data that is passed to the asynchronous method. |
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose and read|

##### Example

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### getCallbackTokenAsync(callback, [userContext])

Gets a string that contains a token used to get an attachment or item from an Exchange Server.

The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.

You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.

In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.|
|`userContext`| Object| &lt;optional&gt;|Any state data that is passed to the asynchronous method.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose and read|

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

The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.<br/><br/>The token is provided as a string in the `asyncResult.value` property.|
|`userContext`| Object| &lt;optional&gt;|Any state data that is passed to the asynchronous method.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

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

> [!NOTE]
> This method is not supported in the following scenarios.
> - In Outlook for iOS or Outlook for Android
> - When the add-in is loaded in a Gmail mailbox
> 
> In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.

The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.

You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.

The XML request must specify UTF-8 encoding.

```xml
<?xml version="1.0" encoding="utf-8"?>
```

Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.

##### Version differences

When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`data`| String||The EWS request.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.<br/><br/>The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.|
|`userContext`| Object| &lt;optional&gt;|Any state data that is passed to the asynchronous method.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Minimum permission level](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[Applicable Outlook mode](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose or read|

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