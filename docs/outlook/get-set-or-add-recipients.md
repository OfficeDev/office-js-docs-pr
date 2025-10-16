---
title: Get, set, or add recipients when composing an appointment or message in Outlook
description: Learn how to get, set, or add recipients to a message or appointment in an Outlook add-in.
ms.date: 10/02/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Get, set, or add recipients when composing an appointment or message in Outlook

Easily identify and manage recipients of a message or appointment with the Office JavaScript API.

In this article, you'll learn how to:

- Get existing recipients from messages and appointments
- Set recipients to replace existing ones
- Add new recipients to messages and appointments

## Understanding recipient properties

Different mail item types support different recipient properties. The following table shows which properties are available for messages and appointments.

| Mail item type | Recipient properties |
|----------------|----------------------|
| **Message** | <ul><li>[item.bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)</li><li>[item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)</li><li>[item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)</li></ul> |
| **Appointment** | <ul><li>[item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)</li><li>[item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)</li></ul>

> [!TIP]
> If your add-in operates on both messages and appointments, we recommend calling [Office.context.mailbox.item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) to help identify between the two mail item types. This way, your add-in can access the appropriate recipient properties.

## Try it out

Explore interactive samples to learn how to manage recipients of a mail item. Install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) then try out the following sample snippets.

- Get to (Message Read)
- Get and set to (Message Compose)
- Get cc (Message Read)
- Get and set cc (Message Compose)
- Get and set bcc (Message Compose)
- Get required attendees (Appointment Attendee)
- Get and set required attendees (Appointment Organizer)
- Get optional attendees (Appointment Attendee)
- Get and set optional attendees (Appointment Organizer)

To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

## Get recipients

This section identifies the different methods to get recipients in read and compose modes.

### Get recipients in read mode

In read mode, you can access recipients from the parent object directly, such as the following example.

```js
Office.context.mailbox.item.cc;
```

The recipients are returned as an array of [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) objects. You can then determine the following information about a recipient from their corresponding `EmailAddressDetails` object.

- Display name
- Email address
- [Recipient type](/javascript/api/outlook/office.mailboxenums.recipienttype)

### Get recipients in compose mode

In compose mode, you must call the [getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) method to access recipients because both the user and your add-in might be modifying recipients simultaneously. This asynchronous approach prevents conflicts and ensures data consistency when multiple processes are working with the same item.

The `getAsync` method requires a callback function that receives the recipients as an array of [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) objects. Each object contains the recipient's display name, email address, and type.

> [!TIP]
> Because the `getAsync` method is asynchronous, if there are subsequent actions that depend on successfully getting the recipients, you should organize your code to run these actions only in the corresponding callback function when the asynchronous call has successfully completed.

The following example displays the email addresses of the recipients in a message or appointment.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        getAllRecipients();
    }
});

// Gets the email addresses of all the recipients of the item being composed.
function getAllRecipients() {
    let toRecipients, ccRecipients, bccRecipients;

    // Verify if the mail item is an appointment or message.
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    } else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }

    // Get the recipients from the To or Required field of the item being composed.
    toRecipients.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }

        // Display the email addresses of the recipients or attendees.
        write(`Recipients in the To or Required field: ${displayAddresses(asyncResult.value)}`);
    });

    // Get the recipients from the Cc or Optional field of the item being composed.
    ccRecipients.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }

        // Display the email addresses of the recipients or attendees.
        write(`Recipients in the Cc or Optional field: ${displayAddresses(asyncResult.value)}`);
    });

    // Get the recipients from the Bcc field of the message being composed, if applicable.
    if (bccRecipients.length > 0) {
        bccRecipients.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }

        // Display the email addresses of the recipients.
        write(`Recipients in the Bcc field: ${displayAddresses(asyncResult.value)}`);
        });
    } else {
        write("Recipients in the Bcc field: None");
    }
}

// Displays the email address of each recipient.
function displayAddresses (recipients) {
    for (let i = 0; i < recipients.length; i++) {
        write(recipients[i].emailAddress);
    }
}

// Writes to a div with id="message" on the page.
function write(message) {
    document.getElementById("message").innerText += message;
}
```

> [!TIP]
> To learn more about asynchronous calls, see [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).

#### Resolved recipients

The `getAsync` method only returns recipients resolved by the Outlook client. A resolved recipient has the following characteristics.

- If the recipient has a saved entry in the sender's address book, Outlook resolves the email address to the recipient's saved display name.
- A Teams meeting status icon appears before the recipient's name or email address.
- A semicolon appears after the recipient's name or email address.
- The recipient's name or email address is underlined or enclosed in a box.

To resolve an email address once it's added to a mail item, the sender must use the <kbd>Tab</kbd> key or select a suggested contact or email address from the auto-complete list.

In Outlook on the web and on Windows (new and classic), if a user creates a new message by selecting a contact's email address link from a contact or profile card, they must first resolve the email address so that it can be included in the results of the `getAsync` call.

## Set recipients

The [setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1)) method replaces all existing recipients with a new list. In the `setAsync` call, you must provide an array as the input argument for the `recipients` parameter in of the of the following formats.

- An array of SMTP address strings. For example, `["user@contoso.com", "team@contoso.com"]`.
- An array of dictionaries, each containing a display name and email address. For example, `[{ displayName: "Megan Bowen", emailAddress: "megan@contoso.com" }]`.
- An array of `EmailAddressDetails` objects, similar to the array returned by the `getAsync` method. For example, `[{ displayName: "Megan Bowen", emailAddress: "megan@contoso.com", recipientType: Office.MailboxEnums.RecipientType.User }]`.

> [!NOTE]
> In Outlook on mobile devices, be mindful of the following:
>
> - The `setAsync` method is supported starting in Version 4.2530.0.
> - The `setAsync` method isn't supported when a user replies from the reply field at the bottom of a message.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        setRecipients();
    }
});

// Sets the recipients of the item being composed.
function setRecipients() {
    let toRecipients, ccRecipients, bccRecipients;

    // Verify if the mail item is an appointment or message.
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    } else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }

    // Set the recipients in the To or Required field of the item being composed.
    toRecipients.setAsync(
        [{
            "displayName": "Graham Durkin", 
            "emailAddress": "graham@contoso.com"
         },
         {
            "displayName": "Donnie Weinberg",
            "emailAddress": "donnie@contoso.com"
         }],
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(asyncResult.error.message);
                return;
            }

            console.log("Successfully set the recipients in the To or Required field.");
            // Run additional operations appropriate to your scenario.
    });

    // Set the recipients in the Cc or Optional field of the item being composed.
    ccRecipients.setAsync(
        [{
            "displayName": "Perry Horning", 
            "emailAddress": "perry@contoso.com"
         },
         {
            "displayName": "Guy Montenegro",
            "emailAddress": "guy@contoso.com"
         }],
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(asyncResult.error.message);
                return;
            }

            console.log("Successfully set the recipients in the Cc or Optional field.");
            // Run additional operations appropriate to your scenario.
    });

    // Set the recipients in the Bcc field of the message being composed.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                "displayName": "Lewis Cate", 
                "emailAddress": "lewis@contoso.com"
            },
            {
                "displayName": "Francisco Stitt",
                "emailAddress": "francisco@contoso.com"
            }],
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log(asyncResult.error.message);
                    return;
                }
    
                console.log("Successfully set the recipients in the Bcc field.");
                // Run additional operations appropriate to your scenario.
        });
    }
}
```

## Add recipients

If you don't want to overwrite any existing recipients in an appointment or message, call the [addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)) method. Similar to the `setAsync` method, the `addAsync` method also requires a `recipients` input argument. You can optionally provide a callback function, and any arguments for the callback using the `asyncContext` parameter.

> [!NOTE]
>
> In Outlook on mobile devices, be mindful of the following:
>
> - The `addAsync` method is supported starting in Version 4.2530.0.
> - The `addAsync` method isn't supported when a user replies from the reply field at the bottom of a message.

The following example checks if the item being composed is an appointment, then appends two required attendees to it.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        addAttendees();
    }
});

// Adds the specified recipients as required attendees of the appointment.
function addAttendees() {
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName": "Kristie Jensen",
            "emailAddress": "kristie@contoso.com"
         },
         {
            "displayName": "Pansy Valenzuela",
            "emailAddress": "pansy@contoso.com"
          }],
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(asyncResult.error.message);
                return;
            }

            console.log("Successfully added the required attendees.");
            // Run additional operations appropriate to your scenario.
        });
    }
}
```

## See also

- [Create Outlook add-ins for compose forms](compose-scenario.md)
- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
- [Get or set the subject when composing an appointment or message in Outlook](get-or-set-the-subject.md)
- [Insert data in the body when composing an appointment or message in Outlook](insert-data-in-the-body.md)
- [Get or set the location when composing an appointment in Outlook](get-or-set-the-location-of-an-appointment.md)
- [Get or set the time when composing an appointment in Outlook](get-or-set-the-time-of-an-appointment.md)
