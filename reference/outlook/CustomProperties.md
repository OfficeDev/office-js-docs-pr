 

# CustomProperties

The `CustomProperties` object represents custom properties that are specific to a particular item and specific to a mail add-in for Outlook. For example, there might be a need for a mail add-in to save some data that is specific to the current email message that activated the add-in. If the user revisits the same message in the future and activates the mail add-in again, the add-in will be able to retrieve the data that had been saved as custom properties.

Because Outlook for Mac doesn’t cache custom properties, if the user’s network goes down, mail add-ins cannot access their custom properties.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

### Example

The following example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the [`saveAsync`](CustomProperties.md#saveAsync) method to save these properties back to the server. After loading the custom properties, the example uses the [`get`](CustomProperties.md#get) method to read the custom property `myProp`, the [`set`](CustomProperties.md#set) method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.

```
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var mailbox = Office.context.mailbox;
    mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
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

### Methods

####  get(name) → {String}

Returns the value of the specified custom property.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The name of the custom property to be returned.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Returns:

The value of the specified custom property.

<dl class="param-type">

<dt>Type</dt>

<dd>String</dd>

</dl>

####  remove(name)

Removes the specified property from the custom property collection.

To make the removal of the property permanent, you must call the [`saveAsync`](CustomProperties.md#saveAsync) method of the `CustomProperties` object.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The name of the property to be removed.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|
####  saveAsync([callback], [asyncContext])

Saves item-specific custom properties to the server.

You must call the `saveAsync` method to persist any changes made with the [`set`](CustomProperties.md#set) method or the [`remove`](CustomProperties.md#remove) method of the `CustomProperties` object. The saving action is asynchronous.

It’s a good practice to have your callback function check for and handle errors from `saveAsync`. In particular, a read add-in can be activated while the user is in a connected state in a read form, and subsequently the user becomes disconnected. If the add-in calls `saveAsync` while in the disconnected state, `saveAsync` would return an error. Your callback method should handle this error accordingly.

##### Parameters:

|Name| Type| Attributes| Description|
|---|---|---|---|
|`callback`| function| &lt;optional&gt;|When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](simple-types.md#asyncresult) object. |
|`asyncContext`| Object| &lt;optional&gt;|Any state data that is passed to the callback method.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

The following JavaScript code sample shows how to asynchronously use the `loadCustomPropertiesAsync` method to load custom properties that are specific to the current item, and the [`saveAsync`](CustomProperties.md#saveAsync) method to save these properties back to the server. After loading the custom properties, the code sample uses the [`get`](CustomProperties.md#get) method to read the custom property `myProp`, the [`set`](CustomProperties.md#set) method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.

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
  if (asyncResult.status == Office.AsyncResultStatus.Failed){
    write(asyncResult.error.message);
  }
  else {
    // Async call to save custom properties completed.
    // Proceed to do the appropriate for your add-in.
  }
}

// Writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message; 
}
```

####  set(name, value)

Sets the specified property to the specified value.

The `set` method sets the specified property to the specified value. You must use the [`saveAsync`](CustomProperties.md#saveAsync) method to save the property to the server.

The `set` method creates a new property if the specified property does not already exist; otherwise, the existing value is replaced with the new value. The `value` parameter can be of any type; however, it is always passed to the server as a string.

##### Parameters:

|Name| Type| Description|
|---|---|---|
|`name`| String|The name of the property to be set.|
|`value`| Object|The value of the property to be set.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| ReadItem|
|Applicable Outlook mode| Compose or read|