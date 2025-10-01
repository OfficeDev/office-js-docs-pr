---
title: Get or set the location of an appointment in an add-in
description: Learn how to get or set the location of an appointment in an Outlook add-in.
ms.date: 10/02/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Get or set the location when composing an appointment in Outlook

Learn how to build an Outlook add-in that effectively manages appointment locations.

## Choose the right API for your scenario

There are two APIs you can use to manage an appointment's locations.

- [item.enhancedLocation](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
- [item.location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)

The following table compares the two location APIs to help you choose the right approach.

| Feature | `enhancedLocation` API | `location` API |
|---------|----------------|------------------------|
| **Minimum requirement set** | [1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) | [1.1](/javascript/api/requirement-sets/outlook/requirement-set-1.1/outlook-requirement-set-1.1) |
| **Recommended use** | Use the [enhancedLocation API](#use-the-enhancedlocation-api) to better identify and manage locations, especially if you need to determine the [location type](/javascript/api/outlook/office.mailboxenums.locationtype). | Use the [location API](#use-the-location-api) if Outlook clients don't support requirement set 1.8 or later or if you only need basic string-based location management. |
| **Supported input and output types** | [Office.LocationIdentifier](/javascript/api/outlook/office.locationidentifier) | String |
| **Supported operations** | <ul><li>Get</li><li>Set</li><li>Remove</li></ul> | <ul><li>Get</li><li>Set</li></ul> |

> [!TIP]
> For guidance on how to check if an Outlook client supports a particular requirement set, see [Use APIs from later requirement sets](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#use-apis-from-later-requirement-sets).

## Try it out

Try interactive samples to see the location APIs in action. Install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) then try out the following sample snippets.

- Get the location (Read) - implements the `location` API
- Get and set the location (Appointment Organizer) - implements the `location` API
- Manage the locations of an appointment - implements the `enhancedLocation` API

To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

## API availability by mode

The following table lists the location APIs and their availability in compose and read modes.

| API | Applicable appointment modes |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | <ul><li>Read (Attendee)</li></ul> |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | <ul><li>Compose (Organizer)</li></ul> |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | <ul><li>Compose (Organizer)</li></ul> |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | <ul><li>Read (Attendee)</li><li>Compose (Organizer)</li></ul> |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | <ul><li>Compose (Organizer)</li></ul> |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | <ul><li>Compose (Organizer)</li></ul> |

## Use the enhancedLocation API

On Outlook clients that support requirement set 1.8 or later, use the `enhancedLocation` API to [add](#add-location), [get](#get-location), and [remove](#remove-location) locations.

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
>
> - [Personal contact groups](https://support.microsoft.com/office/88ff6c60-0a1d-4b54-8c9d-9e1a71bc3023) added as appointment locations aren't returned by the [enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) method.
> - If a location was added using the `item.location` API, its room type is `LocationType.Custom`.

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

## Use the location API

On Outlook clients that don't support requirement set 1.8 or later, use the `location` API to [get](#get-the-location) and [set](#set-the-location) locations. You can also use the `location` API on recent versions of Outlook clients if you don't need advanced features to manage multiple locations.

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
- [Get and set the recurrence of appointments](get-and-set-recurrence.md)
