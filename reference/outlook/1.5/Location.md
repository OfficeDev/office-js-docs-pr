

# Location

Provides methods to get and set the location of a meeting in an Outlook add-in.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose|

##### Members and methods

| Member | Type |
|--------|------|
| [getAsync](#getasyncoptions-callback) | Method |
| [setAsync](#setasynclocation-options-callback) | Method |

### Methods

####  getAsync([options], callback)

Gets the location of an appointment.

The `getAsync` method starts an asynchronous call to the Exchange server to get the location of an appointment.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object.

The location of the appointment is provided as a string in the `asyncResult.value` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose|
####  setAsync(location, [options], [callback])

Sets the location of an appointment.

The `setAsync` method starts an asynchronous call to the Exchange server to set the location of an appointment. Setting the location of an appointment overwrites the current location.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`location`| String||The location of the appointment. The string is limited to 255 characters.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.|
|`options.asyncContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback method.|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>If setting the location fails, the `asyncResult.error` property will contain an error code.|

##### Errors

| Error code | Description |
|------------|-------------|
| `DataExceedsMaximumSize` | The `location` parameter is longer than 255 characters. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose|
