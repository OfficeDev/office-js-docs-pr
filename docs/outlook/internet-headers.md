---
title: Get and set internet headers on a message in an Outlook add-in
description: Learn how to use the internet headers API in Outlook add-ins to persist custom data on email messages across all clients, and how to read those headers from received messages.
ms.date: 05/06/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Get and set internet headers on a message in an Outlook add-in

Internet headers let your Outlook add-in attach custom key-value data to an outgoing message that persists after the message leaves Exchange. Unlike item-level custom properties, internet headers travel with the email, so recipients' add-ins can read them. This article shows how to set internet headers while composing a message and how to read them from a received message.

## Why use internet headers

Outlook add-ins can store custom data at different scopes.

- **Item level**: Use [CustomProperties](/javascript/api/outlook/office.customproperties) for values tied to one item across sessions or [SessionData](/javascript/api/outlook/office.sessiondata) for values needed only during the current compose session.
- **Mailbox level**: Use [RoamingSettings](/javascript/api/outlook/office.roamingsettings) for values that apply across the user's mailbox.

These options don't persist on the message after it leaves Exchange, so recipients can't read those values. Internet headers address this limitation.

Introduced in [Mailbox requirement set 1.8](/javascript/api/requirement-sets/outlook/outlook-requirement-set-1-8), internet headers APIs let you:

- Stamp information on an email that persists after it leaves Exchange across clients.
- Read persisted information from an email in read scenarios across clients.
- Access the full MIME header of the email.

In Exchange on-premises environments, you can set internet headers through Exchange Web Services (EWS) requests. However, some scenarios can fail. For example, in compose mode on Outlook desktop, the item ID isn't synced on `saveAsync` in cached mode.

> [!TIP]
> To learn more about these options, see [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md).

:::image type="content" source="../images/outlook-internet-headers.png" alt-text="Diagram of internet headers. Text: User 1 sends email. Add-in manages custom internet headers while user is composing email. User 2 receives the email. Add-in gets internet headers from received email then parses and uses custom headers.":::

## Supported clients

To use the internet headers API in your add-in, your Outlook client must support requirement set 1.8 or later. For information on supported clients, see [Outlook client support](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support).

The internet headers API is also supported in Outlook on Android and on iOS starting in Version 4.2405.0. To learn more about features supported in Outlook on mobile devices, see [Outlook JavaScript APIs supported in Outlook on mobile devices](outlook-mobile-apis.md).

## Set internet headers while composing a message

Use the [item.internetHeaders](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-internetheaders-member) property to manage the custom internet headers you place on the current message in Compose mode.

### Set, get, and remove custom internet headers example

The following example shows how to set, get, and remove custom internet headers.

```js
// Set custom internet headers.
function setCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "preferred-fruit": "orange", "preferred-vegetable": "broccoli", "best-vegetable": "spinach" },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.getAsync(
    ["preferred-fruit", "preferred-vegetable", "best-vegetable", "nonexistent-header"],
    getCallback
  );
}

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected headers: " + JSON.stringify(asyncResult.value));
  } else {
    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
  }
}

// Remove custom internet headers.
function removeSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.removeAsync(
    ["best-vegetable", "nonexistent-header"],
    removeCallback);
}

function removeCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully removed selected headers");
  } else {
    console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
  }
}

setCustomHeaders();
getSelectedCustomHeaders();
removeSelectedCustomHeaders();
getSelectedCustomHeaders();

/* Sample output:
Successfully set headers
Selected headers: {"best-vegetable":"spinach","preferred-fruit":"orange","preferred-vegetable":"broccoli"}
Successfully removed selected headers
Selected headers: {"preferred-fruit":"orange","preferred-vegetable":"broccoli"}
*/
```

## Get internet headers while reading a message

Call [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getallinternetheadersasync-member(1)) to get internet headers on the current message in Read mode.

### Get sender preferences from current MIME headers example

Building on the example from the previous section, the following code shows how to get the sender's preferences from the current email's MIME headers.

```js
Office.context.mailbox.item.getAllInternetHeadersAsync(getCallback);

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Sender's preferred fruit: " + asyncResult.value.match(/preferred-fruit:.*/gim)[0].slice(17));
    console.log("Sender's preferred vegetable: " + asyncResult.value.match(/preferred-vegetable:.*/gim)[0].slice(21));
  } else {
    console.log("Error getting preferences from header: " + JSON.stringify(asyncResult.error));
  }
}

/* Sample output:
Sender's preferred fruit: orange
Sender's preferred vegetable: broccoli
*/
```

> [!IMPORTANT]
> This sample works for simple cases. For more complex information retrieval (for example, multi-instance headers or folded values as described in [RFC 2822](https://tools.ietf.org/html/rfc2822)), try using an appropriate MIME-parsing library.

## Recommended practices

Currently, internet headers are a finite resource on a user's mailbox. When the quota is exhausted, you can't create any more internet headers on that mailbox, which can result in unexpected behavior from clients that rely on this to function.

Apply the following guidelines when you create internet headers in your add-in.

- Create the minimum number of headers required. The header quota is based on the total size of headers applied to a message. In Exchange Online, the header limit is capped at 256 KB, while in an Exchange on-premises environment, the limit is determined by your organization's administrator. For further information on header limits, see [Exchange Online message limits](/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits) and [Exchange Server message limits](/exchange/mail-flow/message-size-limits).
- Name headers so that you can reuse and update their values later. As such, avoid naming headers in a variable manner (for example, based on user input, timestamp, etc.).

## See also

- [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md)
- [Limits for activation and API usage in Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
