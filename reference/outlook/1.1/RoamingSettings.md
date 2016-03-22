 

# RoamingSettings

The settings created by using the methods of the `RoamingSettings` object are saved per add-in and per user. That is, they are available only to the add-in that created them, and only from the user's mail box in which they are saved.

> While the Outlook Add-in API limits access to these settings to only the add-in that created them, these settings should not be considered secure storage. They can be accessed by Exchange Web Services or Extended MAPI. They should not be used to store sensitive information such as user credentials or security tokens.

The name of a setting is a String, while the value can be a String, Number, Boolean, null, Object, or Array.

The `RoamingSettings` object is accessible via the [`roamingSettings`](Office.context.md#roamingSettings) property in the `Office.context` namespace.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|

### Example

```
// Get the current value of the 'myKey' setting
var value = Office.context.roamingSettings.get('myKey');
// Update the value of the 'myKey' setting
Office.context.roamingSettings.set('myKey', 'Hello World!');
// Persist the change
Office.context.roamingSettings.saveAsync();
```

### Methods

####  get(name) → (nullable) {String|Number|Boolean|Object|Array}

Retrieves the specified setting.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The case-sensitive name of the setting to retrieve.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|

##### Returns:

<dl class="param-type">

<dt>Type</dt>

<dd>String | Number | Boolean | Object | Array</dd>

</dl>

####  remove(name)

Removes the specified setting.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The case-sensitive name of the setting to remove.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|
####  saveAsync([callback])

Saves the settings.

Any settings previously saved by an add-in are loaded when it is initialized, so during the lifetime of the session you can just use the [`set`](RoamingSettings.md#set) and [`get`](RoamingSettings.md#get) methods to work with the in-memory copy of the settings property bag. When you want to persist the settings so that they are available the next time the add-in is used, use the `saveAsync` method.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. |

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|
####  set(name, value)

Sets or creates the specified setting.

The set method creates a new setting of the specified name if it does not already exist, or sets an existing setting of the specified name. The value is stored in the document as the serialized JSON representation of its data type.

A maximum of 2MB is available for the settings of each add-in, and each individual setting is limited to 32KB.

Any changes made to settings using the `set` function will not be saved to the server until the [`saveAsync`](RoamingSettings.md#saveAsync) function is called.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The case-sensitive name of the setting to set or create.|
|`value`| String &#124; Number &#124; Boolean &#124; Object &#124; Array|The value to be stored.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|