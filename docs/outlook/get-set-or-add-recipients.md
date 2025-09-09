---
title: Get, set, or add recipients when composing an appointment or message in Outlook
description: Learn how to get, set, or add recipients to a message or appointment in an Outlook add-in.
ms.date: 09/09/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Get, set, or add recipients when composing an appointment or message in Outlook

The Office JavaScript API provides asynchronous methods ([Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1)), or [Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))) to respectively get, set, or add recipients to a compose form of an appointment or message. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).

Some of the properties that represent recipients in an appointment or message are available for read access in a compose form and in a read form. These properties include  [optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) for appointments, and [cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), and  [to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) for messages.

In a read form, you can access the property directly from the parent object, such as:

```js
Office.context.mailbox.item.cc;
```

But in a compose form, because both the user and your add-in can be inserting or changing a recipient at the same time, you must use the asynchronous method `getAsync` to get these properties, as in the following example.

```js
Office.context.mailbox.item.cc.getAsync(callback);
```

These properties are available for write access in only compose forms, and not read forms.

As with most asynchronous methods in the JavaScript API for Office, `getAsync`, `setAsync`, and `addAsync` take optional input parameters. For more information on how to specify these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).

## Get recipients

This section shows a code sample that gets the recipients of the appointment or message that is being composed, and displays the email addresses of the recipients.

In the Office JavaScript API, because the properties that represent the recipients of an appointment (`optionalAttendees` and `requiredAttendees`) are different from those of a message ([bcc](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-bcc-member), `cc`, and `to`), you should first use the [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) property to identify whether the item being composed is an appointment or message. In compose mode, all these properties of appointments and messages are [Recipients](/javascript/api/outlook/office.recipients) objects, so you can then call the asynchronous method, `Recipients.getAsync`, to get the corresponding recipients.

To use `getAsync`, provide a callback function to check for the status, results, and any error returned by the asynchronous `getAsync` call. The callback function returns an [asyncResult](/javascript/api/office/office.asyncresult) output parameter. Use its `status` and `error` properties to check for the status and any error messages of the asynchronous call, and its `value` property to get the actual recipients. Recipients are represented as an array of [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) objects. You can also provide additional information to the callback function using the optional `asyncContext` parameter in the `getAsync` call.

Note that because the `getAsync` method is asynchronous, if there are subsequent actions that depend on successfully getting the recipients, you should organize your code to start such actions only in the corresponding callback function when the asynchronous call has successfully completed.

> [!IMPORTANT]
> The `getAsync` method only returns recipients resolved by the Outlook client. A resolved recipient has the following characteristics.
>
> - If the recipient has a saved entry in the sender's address book, Outlook resolves the email address to the recipient's saved display name.
> - A Teams meeting status icon appears before the recipient's name or email address.
> - A semicolon appears after the recipient's name or email address.
> - The recipient's name or email address is underlined or enclosed in a box.
>
> To resolve an email address once it's added to a mail item, the sender must use the <kbd>Tab</kbd> key or select a suggested contact or email address from the auto-complete list.
>
> In Outlook on the web and on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic), if a user creates a new message by selecting a contact's email address link from a contact or profile card, they must first resolve the email address so that it can be included in the results of the `getAsync` call.

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
    }
    else {
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

## Set recipients

This section shows a code sample that sets the recipients of the appointment or message that is being composed by the user. Setting recipients overwrites any existing recipients. This example first verifies if the mail item is an appointment or message, so that it can call the asynchronous method, `Recipients.setAsync`, on the appropriate properties that represent recipients of the appointment or message.

When calling `setAsync`, provide an array as the input argument for the `recipients` parameter, in one of the following formats.

- An array of strings that are SMTP addresses.
- An array of dictionaries, each containing a display name and email address, as shown in the following code sample.
- An array of `EmailAddressDetails` objects, similar to the one returned by the `getAsync` method.

You can optionally provide a callback function as an input argument to the `setAsync` method, to make sure any code that depends on successfully setting the recipients would execute only when that happens. If you implement a callback function, use the `status` and `error` properties of the `asyncResult` output parameter to check the status and any error messages of the asynchronous call. To provide additional information to the callback function, use the optional `asyncContext` parameter in the `setAsync` call.

> [!NOTE]
>
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
    }
    else {
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

If you don't want to overwrite any existing recipients in an appointment or message, instead of using `Recipients.setAsync`, use the `Recipients.addAsync` asynchronous method to append recipients. `addAsync` works similarly as `setAsync` in that it requires a `recipients` input argument. You can optionally provide a callback function, and any arguments for the callback using the `asyncContext` parameter. Then, check the status, result, and any error of the asynchronous `addAsync` call using the `asyncResult` output parameter of the callback function. The following example checks if the item being composed is an appointment, then appends two required attendees to it.

> [!NOTE]
>
> In Outlook on mobile devices, be mindful of the following:
>
> - The `addAsync` method is supported starting in Version 4.2530.0.
> - The `addAsync` method isn't supported when a user replies from the reply field at the bottom of a message.

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
