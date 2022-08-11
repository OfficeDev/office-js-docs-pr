---
title: Manage state and settings for an Outlook add-in
description: Learn how to persist add-in state and settings for an Outlook add-in.
ms.date: 07/08/2022
ms.localizationpriority: medium
---

# Manage state and settings for an Outlook add-in

> [!NOTE]
> Please review [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md) in the **Core concepts** section of this documentation before reading this article.

For Outlook add-ins, the Office JavaScript API provides [RoamingSettings](/javascript/api/outlook/office.roamingsettings) and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](/javascript/api/manifest/id) of the add-in that created them.

|**Object**|**Storage location**|
|:-----|:-----|
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|The user's Exchange server mailbox where the add-in is installed. Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client accessing that user's mailbox.<br/><br/> Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.|

## How to save settings in the user's mailbox for Outlook add-ins as roaming settings

An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox. This data is accessible only by that Outlook add-in on behalf of the user running the add-in. The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.

### Loading roaming settings

The following JavaScript code example shows how to load existing roaming settings.

```js
const _settings = Office.context.roamingSettings;
```

### Creating or assigning a roaming setting

Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-set-member(1)) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)) method.

```js
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

The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function. This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method. When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/office/office.asyncresult) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.

### Removing a roaming setting

Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-remove-member(1)) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## How to save settings per item for Outlook add-ins as custom properties

Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.

Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-set-member(1)) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)) method to persist the changes to the item on the Exchange server.

### Custom properties example

The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.

An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.

```js
const property = _customProps.get("propertyName");
```

This example includes the following functions.

|**Function name**|**Description**|
|:-----|:-----|
| `Office.initialize`|Initializes the add-in and loads the custom properties for the current item from the Exchange server.|
| `customPropsCallback`|Gets the custom properties that are returned from the Exchange server and saves it for later use.|
| `updateProperty`|Sets or updates a specific property, and then saves the change to the Exchange server.|
| `removeProperty`|Removes a specific property, and then persists the removal to the Exchange server.|
| `saveCallback`|Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.|

```js
let _mailbox;
let _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready method.
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

### Platform behavior in emails

The following table summarizes saved custom properties behavior in emails for various Outlook clients.

|Scenario|Windows|Web|Mac|
|---|---|---|---|
|New compose|null|null|null|
|Reply, reply all|null|null|null|
|Forward|Loads parent's properties|null|null|
|Sent item from new compose|null|null|null|
|Sent item from reply or reply all|null|null|null|
|Sent item from forward|Removes parent's properties if not saved|null|null|

To handle the situation on Windows:

1. Check for existing properties on initializing your add-in, and keep them or clear them as needed.
1. When setting custom properties, include an additional property to indicate whether the custom properties were added during message read or by Read mode of the add-in. This will help you differentiate if the property was created during compose or inherited from the parent.
1. To check if the user is forwarding an email or replying, you can use [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#outlook-office-messagecompose-getcomposetypeasync-member(1)) (available from requirement set 1.10).

## See also

- [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md)
- [Initialize your Office Add-in](../develop/initialize-add-in.md)
