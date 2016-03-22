 

# NotificationMessages

## NotificationMessages

The `NotificationMessages` object is returned as the [`notificationMessages`](Office.context.mailbox.item.md#notificationmessages-notificationmessages) property of an item.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

### Methods

####  addAsync(key, JSONmessage, [options], [callback])

Adds a notification to an item.

There are a maximum of 5 notifications per message. Setting more will return a `NumberOfNotificationMessagesExceeded` error.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`key`| String||A developer-specified key used to reference this notification message. Developers can use it to modify this message later. It can't be longer than 32 characters.|
|`JSONmessage`| Object||A JSON object that contains the notification message to be added to the item. It consists of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Description</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>Specifies the type of message. If type is <code>ProgressIndicator</code> or <code>ErrorMessage</code>, an icon is automatically supplied and the message is not persistent. Therefore the icon and persistent properties are not valid for these types of messages. Including them will result in an <code>ArgumentException</code>. If type is <code>ProgressIndicator</code>, the developer should remove or replace the progress indicator when the action is complete.</td></tr><tr><td><code>icon</code></td><td>String</td><td>A reference to an icon that is defined in the manifest in the <code>Resource</code> section. It appears in the infobar area. It is only applicable if the type is <code>InformationalMessage</code>. Specifying this parameter for an unsupported type results in an exception.</td></tr><tr><td><code>message</code></td><td>String</td><td>The text of the notification message. Maximum length is 150 characters. If the developer passes in a longer string, an <code>ArgumentOutOfRange</code> exception is thrown.</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>Only applicable when type is <code>InformationalMessage</code>. If <code>true</code>, the message remains until removed by this add-in or dismissed by the user. If <code>false</code>, it is removed when the user navigates to a different item. For error notifications, the message persists until the user sees it once. Specifying this parameter for an unsupported type throws an exception.</td></tr></tbody></table>|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Attributes</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Developers can provide any object they wish to access in the callback method.</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Create three notifications, each with a different key
Office.context.mailbox.item.notificationMessages.addAsync("progress", { 
  type: "progressIndicator", 
  message : "An add-in is processing this message." 
});
Office.context.mailbox.item.notificationMessages.addAsync("information", { 
  type: "informationalMessage", 
  message : "The add-in processed this message.", 
  icon : "iconid", 
  persistent: false 
});
Office.context.mailbox.item.notificationMessages.addAsync("error", { 
  type: "errorMessage", 
  message : "The add-in failed to process this message." 
});
```

####  getAllAsync([options], [callback])

Returns all keys and messages for an item.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Attributes</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Developers can provide any object they wish to access in the callback method.</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. 

On successful completion, the `asyncResult.value` property will contain an array of [`NotificationMessageDetails`](simple-types.md#notificationmessagedetails) objects.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Get all notifications
Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
  if (asyncResult.status != "failed") {
    Office.context.mailbox.item.notificationMessages.replaceAsync( "notifications", { 
      type: "informationalMessage", 
      message : "Found " + asyncResult.value.length + " notifications.", 
      icon : "iconid", 
      persistent: false 
    });
  }
```

####  removeAsync(key, [options], [callback])

Removes a notification message for an item.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`key`| String||The key for the notification message to remove.|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Attributes</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Developers can provide any object they wish to access in the callback method.</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. 

If the key is not found, a `KeyNotFound` error is returned in the `asyncResult.error` property.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Remove a notification
Office.context.mailbox.item.notificationMessages.removeAsync("progress");
```

####  replaceAsync(key, JSONmessage, [options], [callback])

Replaces a notification message that has a given key with another message.

If a notification message with the specified key doesn't exist, `replaceAsync` will add the notification.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`key`| String||The key for the notification message to replace. It can't be longer than 32 characters.|
|`JSONmessage`| Object||A JSON object that contains the new notification message to replace the existing message. It consists of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Description</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>Specifies the type of message. If type is <code>ProgressIndicator</code> or <code>ErrorMessage</code>, an icon is automatically supplied and the message is not persistent. Therefore the icon and persistent properties are not valid for these types of messages. Including them will result in an <code>ArgumentException</code>. If type is <code>ProgressIndicator</code>, the developer should remove or replace the progress indicator when the action is complete.</td></tr><tr><td><code>icon</code></td><td>String</td><td>A reference to an icon that is defined in the manifest in the <code>Resource</code> section. It appears in the infobar area. It is only applicable if the type is <code>InformationalMessage</code>. Specifying this parameter for an unsupported type results in an exception.</td></tr><tr><td><code>message</code></td><td>String</td><td>The text of the notification message. Maximum length is 150 characters. If the developer passes in a longer string, an <code>ArgumentOutOfRange</code> exception is thrown.</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>Only applicable when type is <code>InformationalMessage</code>. If <code>true</code>, the message remains until removed by this add-in or dismissed by the user. If <code>false</code>, it is removed when the user navigates to a different item. For error notifications, the message persists until the user sees it once. Specifying this parameter for an unsupported type throws an exception.</td></tr></tbody></table>|
|`options`| Object| &lt;optional&gt;|An object literal that contains one or more of the following properties.<br/><br/>**Properties**<br/><table class="nested-table"><thead><tr><th>Name</th><th>Type</th><th>Attributes</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;optional&gt;</td><td>Developers can provide any object they wish to access in the callback method.</td></tr></tbody></table>|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Replace a notification with an informational notification
Office.context.mailbox.item.notificationMessages.replaceAsync("progress", { 
  type: "informationalMessage", 
  message : "The message was processed successfully.", 
  icon : "iconid", 
  persistent: false 
});
```
