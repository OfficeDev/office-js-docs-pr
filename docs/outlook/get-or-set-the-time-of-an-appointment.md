---
title: Get or set appointment time in an Outlook add-in
description: Learn how to get or set the start and end time of an appointment in an Outlook add-in.
ms.date: 10/07/2022
ms.topic: how-to
ms.localizationpriority: medium
---

# Get or set the time when composing an appointment in Outlook

The Office JavaScript API provides asynchronous methods ([Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) and [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))) to get and set the start or end time of an appointment that the user is composing. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in XML manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md). Activation rules aren't supported in add-ins that use a [Unified Microsoft 365 manifest (preview)](../develop/json-manifest-overview.md).

The [start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:

```js
item.start
```

and in:

```js
item.end
```

But in a compose form, because both the user and your add-in can be inserting or changing the time at the same time, you must use the asynchronous method **getAsync** to get the start or end time, as shown below:

```js
item.start.getAsync
```

and:

```js
item.end.getAsync
```

As with most asynchronous methods in the Office JavaScript API, **getAsync** and **setAsync** take optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).

## Get the start or end time

This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.

```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
```

To use **item.start.getAsync** or **item.end.getAsync**, provide a callback function that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback function through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the start time as a **Date** object in UTC format using the [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) property.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## Set the start or end time

This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.

Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment.

To use **item.start.setAsync** or **item.end.setAsync**, specify a **Date** value in UTC in the _dateTime_ parameter. If you get a date based on an input by the user on the client, you can use [mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) to convert the value to a **Date** object in UTC. You can provide an optional callback function and any arguments for the callback function in the _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified start or end time string as plain text, overwriting any existing start or end time for that item.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    const startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
                // Do whatever appropriate for your scenario
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

- [Get and set item data in a compose form in Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Get and set Outlook item data in read or compose forms](item-data.md)
- [Create Outlook add-ins for compose forms](compose-scenario.md)
- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
- [Get, set, or add recipients when composing an appointment or message in Outlook](get-set-or-add-recipients.md)  
- [Get or set the subject when composing an appointment or message in Outlook](get-or-set-the-subject.md)
- [Insert data in the body when composing an appointment or message in Outlook](insert-data-in-the-body.md)
- [Get or set the location when composing an appointment in Outlook](get-or-set-the-location-of-an-appointment.md)
