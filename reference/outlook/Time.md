 

# Time

The `Time` object is returned as the [`start`](Office.context.mailbox.item.md#start-datetime) or [`end`](Office.context.mailbox.item.md#end-datetime) property of an appointment in compose mode.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose|

### Methods

####  getAsync([options], callback)

Gets the start or end time of an appointment.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Attributes</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Developers can provide any object they wish to access in the callback method.</td></tr></tbody></table>|
|`callback`| function||When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. 

The date and time is provided as a Date object in the `asyncResult.value` property. The value is in Coordinated Universal Time (UTC). You can convert the UTC time to the local client time by using the [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose|
####  setAsync(dateTime, [options], [callback])

Sets the start or end time of an appointment.

If the `setAsync` method is called on the [`start`](Office.context.mailbox.item.md#start-datetime) property, the [`end`](Office.context.mailbox.item.md#end-datetime) property will be adjusted to maintain the duration of the appointment as previously set. If the `setAsync` method is called on the `end` property, the duration of the appointment will be extended to the new end time.

The time must be in UTC; you can get the correct UTC time by using the [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) method.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`dateTime`| Date||A Date object in Coordinated Universal Time (UTC).|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Attributes</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Developers can provide any object they wish to access in the callback method.</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. <br/>If setting the date and time fails, the `asyncResult.error` property will contain an error code.<br/><table class="nested-table"><thead><tr><th>Error code</th><th>Description</th></tr></thead><tbody><tr><td><code>InvalidEndTime</code></td><td>The appointment end time is before the appointment start time.</td></tr></tbody></table>|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.1|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadWriteItem|
|Applicable Outlook mode| Compose|

##### Example

The following example sets the start time of an appointment.

```js
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
