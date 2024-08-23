---
title: Get or set the location of an appointment in an add-in
description: Learn how to get or set the location of an appointment in an Outlook add-in.
ms.date: 04/12/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Get or set the location when composing an appointment in Outlook

The Office JavaScript API provides properties and methods to manage the location of an appointment that the user is composing. Currently, there are two properties that provide an appointment's location:

- [item.location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties): Basic API that allows you to get and set the location.
- [item.enhancedLocation](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties): Enhanced API that allows you to get and set the location, and includes specifying the [location type](/javascript/api/outlook/office.mailboxenums.locationtype). The type is `LocationType.Custom` if you set the location using `item.location`.

The following table lists the location APIs and the modes (i.e., Compose or Read) where they are available.

| API | Applicable appointment modes |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | Attendee/Read |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | Organizer/Compose |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | Organizer/Compose |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | Organizer/Compose,<br>Attendee/Read |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | Organizer/Compose |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | Organizer/Compose |

To use the methods that are available only to compose add-ins, configure the add-in only manifest to activate the add-in in Organizer/Compose mode. See [Create Outlook add-ins for compose forms](compose-scenario.md) for more details. Activation rules aren't supported in add-ins that use a [Unified manifest for Microsoft 365](../develop/json-manifest-overview.md).

## Use the `enhancedLocation` API

You can use the `enhancedLocation` API to get and set an appointment's location. The location field supports multiple locations and, for each location, you can set the display name, type, and conference room email address (if applicable). See [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) for supported location types.

### Add location

The following example shows how to add a location by calling [addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member).

```js
let item;
const locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### Get location

The following example shows how to get the location by calling [getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-enhancedlocation-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (place) {
        console.log("Display name: " + place.displayName);
        console.log("Type: " + place.locationIdentifier.type);
        if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
            console.log("Email address: " + place.emailAddress);
        }
    });
}
```

> [!NOTE]
> [Personal contact groups](https://support.microsoft.com/office/88ff6c60-0a1d-4b54-8c9d-9e1a71bc3023) added as appointment locations aren't returned by the [enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) method.

### Remove location

The following example shows how to remove the location by calling [removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## Use the `location` API

You can use the `location` API to get and set an appointment's location.

### Get the location

This section shows a code sample that gets the location of the appointment that the user is composing, and displays the location.

To use `item.location.getAsync`, provide a callback function that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback function through the `asyncContext` optional parameter. You can obtain status, results, and any error using the output parameter `asyncResult` of the callback. If the asynchronous call is successful, you can get the location as a string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) property.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### Set the location

This section shows a code sample that sets the location of the appointment that the user is composing.

To use `item.location.setAsync`, specify a string of up to 255 characters in the data parameter. Optionally, you can provide a callback function and any arguments for the callback function in the `asyncContext` parameter. You should check the status, result, and any error message in the `asyncResult` output parameter of the callback. If the asynchronous call is successful, `setAsync` inserts the specified location string as plain text, overwriting any existing location for that item.

> [!NOTE]
> You can set multiple locations by using a semi-colon as the separator (e.g., 'Conference room A; Conference room B').

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever is appropriate for your scenario,
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## See also

- [Create your first Outlook add-in](../quickstarts/outlook-quickstart-yo.md)
- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
