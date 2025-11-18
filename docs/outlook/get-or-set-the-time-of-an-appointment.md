---
title: Get or set the time when composing an appointment in Outlook
description: Learn how to get or set the start and end time of an appointment in an Outlook add-in.
ms.date: 10/27/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Get or set the time when composing an appointment in Outlook

Programmatically manage the start and end time of appointments to streamline meeting scheduling and enhance productivity in your Outlook add-ins.

The Office JavaScript API provides asynchronous methods ([Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) and [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))) to get and set the start or end time of an appointment being composed. These asynchronous methods are available only to compose add-ins. To use these methods, make sure you have set up the add-in only manifest of the add-in appropriately for Outlook to activate the add-in in compose forms.

The [start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are available for appointments in both compose and read forms. In read form, you can access the properties directly from the parent object, as in:

```js
Office.context.mailbox.item.start;
Office.context.mailbox.item.end;
```

But in a compose form, because both the user and your add-in can be inserting or changing the time at the same time, you must use the `getAsync` asynchronous method to get the start or end time.

```js
Office.context.mailbox.item.start.getAsync(callback);
Office.context.mailbox.item.end.getAsync(callback);
```

As with most asynchronous methods in the Office JavaScript API, `getAsync` and `setAsync` take optional input parameters. For more information on how to specify these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).

## Try it out

Try interactive samples to manage the start and end times of an appointment. Install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) then try out the following sample snippets.

- Get the start date and time (Read)
- Get the end date and time (Read) 
- Get and set the start date and time (Appointment Organizer)
- Get and set the end date and time (Appointment Organizer)

To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

## Get the start or end time

This section shows a code sample that gets the start time of the appointment being composed and displays the time. You can use the same code, but replace the `start` property with the `end` property to get the end time.

To use the `item.start.getAsync` or `item.end.getAsync` methods, provide a callback function that checks the status and result of the asynchronous call. Obtain the status, results, and any error using the [asyncResult](/javascript/api/office/office.asyncresult) output parameter of the callback. If the asynchronous call is successful, use the `asyncResult.value` property to get the start time as a `Date` object in UTC format. To provide any necessary arguments to the callback function, use the `asyncContext` optional parameter of the `getAsync` call.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        getStartTime();
    }
});

// Gets the start time of the appointment being composed.
function getStartTime() {
    item.start.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }

        const startTime = asyncResult.value;
        // Display the start time in UTC format on the page.
        write(`The start time in UTC is: ${startTime.toUTCString()}`);
        // Display the start time in local time on the page.
        write(`The start time in local time is: ${startTime.toLocaleString()}`);
    });
}

// Writes to a div with id="message" on the page.
function write(message) {
    document.getElementById("message").innerText += message;
}
```

## Set the start or end time

This section shows a code sample that sets the start time of an appointment being composed. You can use the same code, but replace the `start` property with the `end` property to set the end time.

To use `item.start.setAsync` or `item.end.setAsync`, specify a UTC-formatted `Date` object in the `dateTime` parameter. If you get a date based on an input by the user in the client, you can use [mailbox.convertToUtcClientTime](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-converttoutcclienttime-member(1)) to convert the value to a `Date` object in the UTC format. If you provide an optional callback function, include the `asyncContext` parameter and add any arguments to it. Additionally, check the status, result, and any error message through the `asyncResult` output parameter of the callback. If the asynchronous call is successful, `setAsync` inserts the specified start or end time string as plain text, overwriting any existing start or end time for that item.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        setStartTime();
    }
});

// Sets the start time of the appointment being composed.
function setStartTime() {
    // Get the current date and time, then add two days to the date.
    const startDate = new Date();
    startDate.setDate(startDate.getDate() + 2);

    item.start.setAsync(
        startDate,
        { asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(asyncResult.error.message);
                return;
            }

            console.log("Successfully set the start time.");
            /*
                Run additional operations appropriate to your scenario and
                use the optionalVariable1 and optionalVariable2 values as needed.
            */
        });
}
```

## Understanding time property interactions and limitations

When working with appointment times, be aware of how changes to the `start` or `end` properties interact with other appointment properties and platform-specific limitations.

- If the appointment being composed already has an existing start time, setting the start time subsequently adjusts the end time to maintain any previous duration of the appointment.
- If the appointment being composed already has an existing end time, setting the end time subsequently adjusts both the duration and end time.
- If the appointment has been set as an all-day event, setting the start time adjusts the end time to 24 hours later, and clears the checkbox for the all-day event in the appointment.
- In classic Outlook on Windows, the `setAsync` method can't be used to change the start or end time of a recurring appointment.

## See also

- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
- [Get or set the location when composing an appointment in Outlook](get-or-set-the-location-of-an-appointment.md)
- [Get, set, or add recipients to an appointment or message in Outlook](get-set-or-add-recipients.md)
- [Get or set the subject when composing an appointment or message in Outlook](get-or-set-the-subject.md)
- [Insert data in the body when composing an appointment or message in Outlook](insert-data-in-the-body.md)
