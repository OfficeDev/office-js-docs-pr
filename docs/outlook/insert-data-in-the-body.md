---
title: Get or set the body of a message or appointment in Outlook
description: Learn how to get or insert data into the body of an appointment or message of an Outlook add-in.
ms.date: 06/03/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Get or set the body of a message or appointment in Outlook

Call the [Body](/javascript/api/outlook/office.body) API on a message or appointment to retrieve content, determine its format, or update content. With the available Body methods, you can customize signatures depending on mail item recipients or add disclaimers for legal purposes.

Select the applicable tab to learn how to get or set the body of a mail item.

# [Get body](#tab/get)

You can get the body of a message or appointment in both read and compose modes. To retrieve the body of a mail item, call [Office.context.mailbox.item.body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1)). When you call the `getAsync` method, you must specify the format for the returned body in the `coercionType` parameter. For example, you can get the body in HTML or plain text format.

The following example gets the body of an item in HTML format.

```javascript
// Get the current body of the message or appointment.
Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
  if (bodyResult.status === Office.AsyncResultStatus.Failed) {
    console.log(`Failed to get body: ${bodyResult.error.message}`);
    return;
  }

  const body = bodyResult.value;

  // Perform additional operations here.
});
```

## Get the body of message replies in Outlook on the web or the new Outlook on Windows

In Outlook on the web and the [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), users can organize their messages as conversations or individual messages in **Settings** > **Mail** > **Layout** > **Message organization**. This setting affects how much of a message's body is displayed to the user, particularly in conversation threads with multiple messages. Depending on the setting, the contents of the entire conversation thread or just the current message is displayed. For more information on the **Message Organization** setting, see [Change how the message list is displayed in Outlook](https://support.microsoft.com/office/57fe0cd8-e90b-4b1b-91e4-a0ba658c0042).

When you call `Office.context.mailbox.item.body.getAsync` on a message reply, the entire body of a conversation thread is returned. If you want the returned body to reflect the user's **Message Organization** setting, you can specify the [bodyMode](/javascript/api/outlook/office.mailboxenums.bodymode) option in the `getAsync` call. The following table lists the portion of the body returned depending on the `bodyMode` configuration.

| bodyMode configuration | Effect on body |
| ----- | ----- |
| `bodyMode` isn't specified in the `getAsync` call | The entire body of the conversation thread is returned. |
| `bodyMode` is set to `Office.MailboxEnums.BodyMode.FullBody` | The entire body of the conversation thread is returned. |
| `bodyMode` is set to `Office.MailboxEnums.BodyMode.HostConfig` | If **Message Organization** is set to **Group messages by conversation** > **All messages from the selected conversation** or **Show email grouped by conversation** > **Newest on top**/**Newest on bottom**, only the body of the current reply is returned.<br><br>If **Message Organization** is set to **Individual messages: Do not group messages** > **Only a single message** or **Show email as individual messages**, the entire body of the conversation thread is returned. |

> [!NOTE]
> The `bodyMode` option is ignored in Outlook on Windows (classic), on Mac, and on mobile devices.

The following example specifies the `bodyMode` option to honor the user's message setting.

```javascript
Office.context.mailbox.item.body.getAsync(
  Office.CoercionType.Html,
  { bodyMode: Office.MailboxEnums.BodyMode.HostConfig },
  (bodyResult) => {
    if (bodyResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Failed to get body: ${bodyResult.error.message}`);
      return;
    }

    const body = bodyResult.value;

    // Perform additional operations here.
  }
);
```

# [Set body](#tab/set)

Use the asynchronous methods ([Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1)), [Body.getTypeAsync](/javascript/api/outlook/office.body#outlook-office-body-gettypeasync-member(1)), [Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)) and [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))) to get the body type then insert data in the body of an appointment or message being composed. These asynchronous methods are only available to compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately so that Outlook activates your add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).

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

## Set the body of message replies in Outlook on the web or the new Outlook on Windows

In Outlook on the web and the [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), users can organize their messages as conversations or individual messages in **Settings** > **Mail** > **Layout** > **Message organization**. This setting affects how much of a message's body is displayed to the user, particularly in conversation threads with multiple messages. Depending on the setting, the contents of the entire conversation thread or just the current message is displayed. For more information on the **Message Organization** setting, see [Change how the message list is displayed in Outlook](https://support.microsoft.com/office/57fe0cd8-e90b-4b1b-91e4-a0ba658c0042).

When you call `Office.context.mailbox.item.body.setAsync` on a message reply, the entire body of a conversation thread is replaced with the text you specify. If you want to honor the user's **Message Organization** setting and only replace the body of the current reply, you can specify the [bodyMode](/javascript/api/outlook/office.mailboxenums.bodymode) option in the `setAsync` call. The following table lists the `bodyMode` configurations and how each affects the message body being set.

| bodyMode configuration | Effect on body |
| ----- | ----- |
| `bodyMode` isn't specified in the `setAsync` call | The entire body of the conversation thread is replaced. This applies even if a user's messages are organized by conversation. In this scenario, the user's setting is temporarily changed to **Individual messages: Do not group messages** > **Only a single message** or **Show email as individual messages** during the `setAsync` call. A notification is shown to the user to alert them to this change. Once the call completes, the user's setting is reinstated. |
| `bodyMode` is set to `Office.MailboxEnums.BodyMode.FullBody` | The entire body of the conversation thread is replaced. This applies even if a user's messages are organized by conversation. In this scenario, the user's setting is temporarily changed to **Individual messages: Do not group messages** > **Only a single message** or **Show email as individual messages** during the `setAsync` call. A notification is shown to the user to alert them to this change. Once the call completes, the user's setting is reinstated. |
| `bodyMode` is set to `Office.MailboxEnums.BodyMode.HostConfig` | If **Message Organization** is set to **Group messages by conversation** > **All messages from the selected conversation** or **Show email grouped by conversation** > **Newest on top**/**Newest on bottom**, only the body of the current reply is replaced.<br><br>If **Message Organization** is set to **Individual messages: Do not group messages** > **Only a single message** or **Show email as individual messages**, the entire body of the conversation thread is replaced. |

> [!NOTE]
> The `bodyMode` option is ignored in Outlook on Windows (classic), on Mac, and on mobile devices.

The following example specifies the `bodyMode` option to honor the user's message setting.

```javascript
Office.context.mailbox.item.body.setAsync(
  "This text replaces the body of the message.",
  {
    coercionType: Office.CoercionType.Html,
    bodyMode: Office.MailboxEnums.BodyMode.HostConfig
  },
  (bodyResult) => {
    if (bodyResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Failed to set body: ${bodyResult.error.message}`);
      return;
    }

    console.log("Successfully replaced the body of the message.");
  }
);
```

---

## Try code samples in Script Lab

Get the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) and try out the item body code samples to see the get and set APIs in action. To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

## See also

- [Create Outlook add-ins for compose forms](compose-scenario.md)
- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
- [Prepend or append content to a message or appointment body on send](append-on-send.md)
- [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Get, set, or add recipients when composing an appointment or message in Outlook](get-set-or-add-recipients.md)
- [Get or set the subject when composing an appointment or message in Outlook](get-or-set-the-subject.md)
