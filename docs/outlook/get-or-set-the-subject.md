---
title: Get or set the subject in an Outlook add-in
description: Learn how to get or set the subject of a message or appointment in an Outlook add-in.
ms.date: 08/09/2023
ms.topic: how-to
ms.localizationpriority: medium
---

# Get or set the subject when composing an appointment or message in Outlook

The Office JavaScript API provides asynchronous methods ([subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1)) and [subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))) to get and set the subject of an appointment or message that the user is composing. These asynchronous methods are available only to compose add-ins. To use these methods, make sure you have set up the add-in only manifest appropriately for Outlook to [activate the add-in in compose forms](compose-scenario.md).

The `subject` property is available for read access in both compose and read forms of appointments and messages. In a read form, access the property directly from the parent object, as in:

```js
Office.context.mailbox.item.subject;
```

But in a compose form, because both the user and your add-in can be inserting or changing the subject at the same time, you must use the `getAsync` method to get the subject asynchronously.

```js
Office.context.mailbox.item.subject.getAsync(callback);
```

The `subject` property is available for write access in only compose forms and not in read forms.

> [!TIP]
> To temporarily set the content displayed in the subject of a message in read mode, use [Office.context.mailbox.item.display.subject (preview)](/javascript/api/outlook/office.display?view=outlook-js-preview&preserve-view=true#outlook-office-display-subject-member).

As with most asynchronous methods in the Office JavaScript API, `getAsync` and `setAsync` take optional input parameters. For more information on how to specify these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).

## Get the subject

This section shows a code sample that gets the subject of the appointment or message that the user is composing, and displays the subject.

To use `item.subject.getAsync`, provide a callback function that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback function through the optional `asyncContext` parameter. To obtain the status, results, and any error from the callback function, use the `asyncResult` output parameter of the callback. If the asynchronous call is successful, use the [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) property to get the subject as a plain text string.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        getSubject();
    }
});

// Gets the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }

        // Display the subject on the page.
        write(`The subject is: ${asyncResult.value}`);
    });
}


// Writes to a div with id="message" on the page.
function write(message) {
    document.getElementById("message").innerText += message; 
}
```

## Set the subject

This section shows a code sample that sets the subject of the appointment or message that the user is composing.

To use `item.subject.setAsync`, specify a string of up to 255 characters in the `data` parameter. Optionally, you can provide a callback function and any arguments for the callback function in the `asyncContext` parameter. Check the callback status, result, and any error message in the `asyncResult` output parameter of the callback. If the asynchronous call is successful, `setAsync` inserts the specified subject string as plain text, overwriting any existing subject for that item.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        setSubject();
    }
});

// Sets the subject of the item that the user is composing.
function setSubject() {
    // Customize the subject with today's date.
    const today = new Date();
    const subject = `Summary for ${today.toLocaleDateString()}`;

    item.subject.setAsync(
        subject,
        { asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
                return;
            }

            /*
              The subject was successfully set.
              Run additional operations appropriate to your scenario and
              use the optionalVariable1 and optionalVariable2 values as needed.
            */
        });
}

// Writes to a div with id="message" on the page.
function write(message) {
    document.getElementById("message").innerText += message; 
}
```

## See also

- [Create Outlook add-ins for compose forms](compose-scenario.md)
- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
- [Get, set, or add recipients when composing an appointment or message in Outlook](get-set-or-add-recipients.md)  
- [Insert data in the body when composing an appointment or message in Outlook](insert-data-in-the-body.md)
- [Get or set the location when composing an appointment in Outlook](get-or-set-the-location-of-an-appointment.md)
- [Get or set the time when composing an appointment in Outlook](get-or-set-the-time-of-an-appointment.md)
