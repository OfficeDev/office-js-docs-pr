
# Persisting add-in state and settings

Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location.

To do that, you can:


- Use members of the JavaScript API for Office that store data as name/value pairs in a property bag stored in a location that depends on add-in type.
    
- Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](http://msdn.microsoft.com/en-us/library/cc848902%28v=vs.85%29.aspx) or [sessionStorage](http://msdn.microsoft.com/en-us/library/cc197020%28v=vs.85%29.aspx)).
    
This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).

## Persisting add-in state and settings with the JavaScript API for Office


The JavaScript API for Office provides the [Settings](../../reference/shared/document.settings.md), [RoamingSettings](../../reference/outlook/RoamingSettings.md), and [CustomProperties](../../reference/outlook/CustomProperties.md) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx) of the add-in that created them.



|**Object**|**Add-in type support**|**Storage location**|**Office host support**|
|:-----|:-----|:-----|:-----|
|[Settings](../../reference/shared/document.settings.md)|content and task pane|The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved. **Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|Word, Excel, or PowerPoint **Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](../../reference/outlook/RoamingSettings.md)|Outlook|The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox. Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.|Outlook|
|[CustomProperties](../../reference/outlook/CustomProperties.md)|Outlook|The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.|Outlook|

## Settings data is managed in memory at runtime


Internally, the data in the property bag accessed with the  **Settings**,  **CustomProperties**, or  **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs. The name (key) for each value must be a **string**, and the stored value can be a JavaScript  **string**,  **number**,  **date**, or  **object**, but not a  **function**.

This example of the property bag structure contains three defined  **string** values named `firstName`,  `location`, and  `defaultView`.




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session. During the session, the settings are managed in entirely in memory using the  **get**,  **set**, and  **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**,  **CustomProperties**, or  **RoamingSettings**). 


 >**Important**  To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the  **saveAsync** method of the corresponding object used to work with that kind of settings. The **get**,  **set**, and  **remove** methods operate only on the in-memory copy of the settings property bag. If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost. 


## How to save add-in state and settings per document for content and task pane add-ins


To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](../../reference/shared/document.settings.md) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.

The  **Settings** object is automatically loaded as part of the [Document](../../reference/shared/document.md) object, and is available when the task pane or content add-in is activated. After the **Document** object is instantiated, you can access the **Settings** object with the [settings](../../reference/shared/document.settings.md) property of the **Document** object. During the lifetime of the session, you can just use the **Settings.get**,  **Settings.set**, and  **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.

Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](../../reference/shared/settings.saveasync.md) method.


### Creating or updating a setting value

The following code example shows how to use the [Settings.set](../../reference/shared/settings.set.md) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.


```
Office.context.document.settings.set('themeColor', 'green');
```

 The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.


### Getting the value of a setting

The following example shows how use the [Settings.get](../../reference/shared/settings.get.md) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.


### Removing a setting

The following example shows how to use the [Settings.remove](../../reference/shared/settings.removehandlerasync.md) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.


```
Office.context.document.settings.remove('themeColor');
```

Nothing will happen if the setting does not exist. Use the  **Settings.saveAsync** method to persist removal of the setting from the document.


### Saving your settings

To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](../../reference/shared/settings.saveasync.md) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter. 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

The anonymous function passed into the  **saveAsync** method as the _callback_ parameter is executed when the operation is completed. The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation. In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.


## How to save settings in the user's mailbox for Outlook add-ins as roaming settings


An Outlook add-in can use the [RoamingSettings](../../reference/outlook/RoamingSettings.md) object to save add-in state and settings data that is specific to the user's mailbox. This data is accessible only by that Outlook add-in on behalf of the user running the add-in. The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into his or her account and runs the Outlook add-in.


### Loading roaming settings


An Outlook add-in typically loads roaming settings in the [Office.initialize](../../reference/shared/office.initialize.md) event handler. The following JavaScript code example shows how to load existing roaming settings.


```
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}

```


### Creating or assigning a roaming setting


Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) method.


```
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

The  **saveAsync** method saves roaming settings asynchronously and takes an optional callback function. This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method. When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](../../reference/outlook/simple-types.md) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.


### Removing a roaming setting


Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## How to save settings per item for Outlook add-ins as custom properties


Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.

Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](../../reference/outlook/CustomProperties.md) and [get](../../reference/outlook/RoamingSettings.md) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](../../reference/outlook/CustomProperties.md) method to persist the changes to the item on the Exchange server.


### Custom properties example

The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties. 

An Outlook add-in that uses these functions retrieves any custom properties by calling the  **get** method on the `_customProps` variable, as shown in the following example.




```
var property = _customProps.get("propertyName");
```

This example includes the following functions:



|**Function name**|**Description**|
|:-----|:-----|
| `Office.initialize`|Initializes the add-in and loads the custom properties for the current item from the Exchange server.|
| `customPropsCallback`|Gets the custom properties that are returned from the Exchange server and saves it for later use.|
| `updateProperty`|Sets or updates a specific property, and then saves the change to the Exchange server.|
| `removeProperty`|Removes a specific property, and then persists the removal to the Exchange server.|
| `saveCallback`|Callback for calls to the  **saveAsync** method in the `updateProperty` and `removeProperty` functions.|



```
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change 
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal 
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method. 
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```


## Additional resources



- [Understanding the JavaScript API for Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
