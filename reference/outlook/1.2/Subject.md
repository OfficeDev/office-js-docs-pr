 

# Subject

Provides methods to get and set the subject of an appointment or message in an Outlook add-in.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose|

### Methods

####  getAsync([options], callback)

Gets the subject of an appointment or message.

The `getAsync` method starts an asynchronous call to the Exchange server to get the subject of an appointment or message.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Attributes</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Developers can provide any object they wish to access in the callback method.</td></tr></tbody></table>|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. 

The subject of the item is provided as a string in the `asyncResult.value` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose|
####  setAsync(subject, [options], [callback])

Sets the subject of an appointment or message.

The `setAsync` method starts an asynchronous call to the Exchange server to set the subject of an appointment or message. Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`subject`| String||The subject of the appointment or message. The string is limited to 255 characters.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Attributes</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Developers can provide any object they wish to access in the callback method.</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>If setting the subject fails, the `asyncResult.error` property will contain an error code.<br/><table class="nested-table"><thead><tr><th>Error code</th><th>Description</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>The <code>subject</code> parameter is longer than 255 characters.</td></tr></tbody></table>|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose|
