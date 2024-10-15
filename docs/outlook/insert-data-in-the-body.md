---
title: Insert data in the body when composing an appointment or message in Outlook
description: Learn how to insert data into the body of an appointment or message in an Outlook add-in.
ms.date: 08/09/2023
ms.topic: how-to
ms.localizationpriority: medium
---

# Insert data in the body when composing an appointment or message in Outlook

Use the asynchronous methods ([Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1)), [Body.getTypeAsync](/javascript/api/outlook/office.body#outlook-office-body-gettypeasync-member(1)), [Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)) and [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))) to get the body type and insert data in the body of an appointment or message being composed. These asynchronous methods are only available to compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately so that Outlook activates your add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).

In Outlook, a user can create a message in text, HTML, or Rich Text Format (RTF), and can create an appointment in HTML format. Before inserting data, you must first verify the supported item format by calling `getTypeAsync`, as you may need to take additional steps. The value that `getTypeAsync` returns depends on the original item format, as well as the support of the device operating system and application to edit in HTML format. Once you've verified the item format, set the `coercionType` parameter of `prependAsync` or `setSelectedDataAsync` accordingly to insert the data, as shown in the following table. If you don't specify an argument, `prependAsync` and `setSelectedDataAsync` assume the data to insert is in text format.

|Data to insert|Item format returned by getTypeAsync|coercionType to use|
|:-----|:-----|:-----|
|Text|Text<sup>1</sup>|Text|
|HTML|Text<sup>1</sup>|Text<sup>2</sup>|
|Text|HTML|Text/HTML|
|HTML|HTML |HTML|

> [!NOTE]
> <sup>1</sup> On tablets and smartphones, `getTypeAsync` returns "Text" if the operating system or application doesn't support editing an item, which was originally created in HTML, in HTML format.
>
> <sup>2</sup> If your data to insert is HTML and `getTypeAsync` returns a text type for the current mail item, you must reorganize your data as text and set `coercionType` to `Office.CoercionType.Text`. If you simply insert the HTML data into a text-formatted item, the application displays the HTML tags as text. If you attempt to insert the HTML data and set `coercionType` to `Office.CoercionType.Html`, you'll get an error.

In addition to the `coercionType` parameter, as with most asynchronous methods in the Office JavaScript API, `getTypeAsync`, `prependAsync`, and `setSelectedDataAsync` take other optional input parameters. For more information on how to specify these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).

## Insert data at the current cursor position

This section shows a code sample that uses `getTypeAsync` to verify the body type of the item that is being composed, and then uses `setSelectedDataAsync` to insert data at the current cursor location.

You must pass a data string as an input parameter to `setSelectedDataAsync`. Depending on the type of the item body, you can specify this data string in text or HTML format accordingly. As mentioned earlier, you can optionally specify the type of the data to be inserted in the `coercionType` parameter. To get the status and results of `setSelectedDataAsync`, pass a callback function and optional input parameters to the method, then extract the needed information from the [asyncResult](/javascript/api/office/office.asyncresult) output parameter of the callback. If the method succeeds, you can get the type of the item body from the `asyncResult.value` property, which is either "text" or "html".

If the user hasn't placed the cursor in the item body, `setSelectedDataAsync` inserts the data at the top of the body. If the user has selected text in the item body, `setSelectedDataAsync` replaces the selected text with the data you specify. Note that `setSelectedDataAsync` can fail if the user simultaneously changes the cursor position while composing the item. The maximum number of characters you can insert at one time is 1,000,000 characters.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        setItemBody();
    }
});

// Inserts data at the current cursor position.
function setItemBody() {
    // Identify the body type of the mail item.
    item.body.getTypeAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
        }

        // Insert data of the appropriate type into the body.
        if (asyncResult.value === Office.CoercionType.Html) {
            // Insert HTML into the body.
            item.body.setSelectedDataAsync(
                "<b> Kindly note we now open 7 days a week.</b>",
                { coercionType: Office.CoercionType.Html, asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                        return;
                    }

                    /*
                      Run additional operations appropriate to your scenario and
                      use the optionalVariable1 and optionalVariable2 values as needed.
                    */
            });
        }
        else {
            // Insert plain text into the body.
            item.body.setSelectedDataAsync(
                "Kindly note we now open 7 days a week.",
                { coercionType: Office.CoercionType.Text, asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                        return;
                    }

                    /*
                      Run additional operations appropriate to your scenario and
                      use the optionalVariable1 and optionalVariable2 values as needed.
                    */
            });
        }
    });
}
```

## Insert data at the beginning of the item body

Alternatively, you can use `prependAsync` to insert data at the beginning of the item body and disregard the current cursor location. Other than the point of insertion, `prependAsync` and `setSelectedDataAsync` behave in similar ways. You must first check the type of the message body to avoid prepending HTML data to a message in text format. Then, pass the data string to be prepended in either text or HTML format to `prependAsync`. The maximum number of characters you can prepend at one time is 1,000,000 characters.

The following JavaScript code first calls `getTypeAsync` to verify the type of the item body. Then, depending on the type, it inserts the data as HTML or text to the top of the body.

```js
let item;

// Confirms that the Office.js library is loaded.
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        item = Office.context.mailbox.item;
        prependItemBody();
    }
});


// Prepends data to the body of the item being composed.
function prependItemBody() {
    // Identify the body type of the mail item.
    item.body.getTypeAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
        }

        // Prepend data of the appropriate type to the body.
        if (asyncResult.value === Office.CoercionType.Html) {
            // Prepend HTML to the body.
            item.body.prependAsync(
                '<b>Greetings!</b>',
                { coercionType: Office.CoercionType.Html, asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                        return;
                    }

                    /*
                      Run additional operations appropriate to your scenario and
                      use the optionalVariable1 and optionalVariable2 values as needed.
                    */
            });
        }
        else {
            // Prepend plain text to the body.
            item.body.prependAsync(
                'Greetings!',
                { coercionType: Office.CoercionType.Text, asyncContext: { optionalVariable1: 1, optionalVariable2: 2 } },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                        return;
                    }

                    /*
                      Run additional operations appropriate to your scenario and
                      use the optionalVariable1 and optionalVariable2 values as needed.
                    */
            });
        }
    });
}
```

## See also

- [Create Outlook add-ins for compose forms](compose-scenario.md)
- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
- [Get, set, or add recipients when composing an appointment or message in Outlook](get-set-or-add-recipients.md)  
- [Get or set the subject when composing an appointment or message in Outlook](get-or-set-the-subject.md)  
- [Get or set the location when composing an appointment in Outlook](get-or-set-the-location-of-an-appointment.md)
- [Get or set the time when composing an appointment in Outlook](get-or-set-the-time-of-an-appointment.md)
