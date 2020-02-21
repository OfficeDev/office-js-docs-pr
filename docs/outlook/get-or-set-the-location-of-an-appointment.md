---
title: Get or set the location of an appointment in an add-in
description: Learn how to get or set the location of an appointment in an Outlook add-in.
ms.date: 10/31/2019
localization_priority: Normal
---

# Get or set the location when composing an appointment in Outlook

The JavaScript API for Office provides properties and methods to manage the location of an appointment that the user is composing. Currently, there are two properties that provide an appointment's location:

- [item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Basic API that allows you to get and set the location.
- [item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Enhanced API that allows you to get and set the location, and includes specifying the [location type](/javascript/api/outlook/office.mailboxenums.locationtype). The type is `LocationType.Custom` if you set the location using `item.location`.

The following table lists the location APIs and the modes (i.e., Compose or Read) where they are available.

| API | Applicable appointment modes |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#location) | Attendee/Read |
| [item.location.getAsync](/javascript/api/outlook/office.location#getasync-options--callback-) | Organizer/Compose |
| [item.location.setAsync](/javascript/api/outlook/office.location#setasync-location--options--callback-) | Organizer/Compose |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | Organizer/Compose,<br>Attendee/Read |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | Organizer/Compose |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | Organizer/Compose |

To use the methods that are available only to compose add-ins, configure the add-in manifest to activate the add-in in Organizer/Compose mode. See [Create Outlook add-ins for compose forms](compose-scenario.md) for more details.

## Use the `enhancedLocation` API

You can use the `enhancedLocation` API to get and set an appointment's location. The location field supports multiple locations and, for each location, you can set the display name, type, and conference room email address (if applicable). See [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) for supported location types.

### Add location

The following example shows how to add a location by calling [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).

```js
var item;
var locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### Get location

The following example shows how to get the location by calling [getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
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

### Remove location

The following example shows how to remove the location by calling [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        Office.context.mailbox.item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## Use the `location` API

You can use the `location` API to get and set an appointment's location.

### Get the location

This section shows a code sample that gets the location of the appointment that the user is composing, and displays the location.

To use `item.location.getAsync`, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the `asyncContext` optional parameter. You can obtain status, results, and any error using the output parameter `asyncResult` of the callback. If the asynchronous call is successful, you can get the location as a string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
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

To use `item.location.setAsync`, specify a string of up to 255 characters in the data parameter. Optionally, you can provide a callback method and any arguments for the callback method in the `asyncContext` parameter. You should check the status, result, and any error message in the `asyncResult` output parameter of the callback. If the asynchronous call is successful, `setAsync` inserts the specified location string as plain text, overwriting any existing location for that item.

> [!NOTE]
> You can set multiple locations by using a semi-colon as the separator (e.g., 'Conference room A; Conference room B').

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
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

- [Create your first Outlook add-in](../quickstarts/outlook-quickstart.md)
- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
