---
title: Manage the delivery date and time of a message
description: Learn how to get and set the delivery date and time of a message in compose mode.
ms.date: 04/12/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Manage the delivery date and time of a message

The Outlook client gives you the option to delay the delivery of a message, but requires you to keep Outlook and your device running to send it at the specified time. With the Office JavaScript API, you can now implement an Outlook add-in that sends scheduled messages even with your Outlook client closed or with your device turned off. This capability provides your users with the convenience to schedule email marketing campaigns or time a message to be delivered during a colleague or customer's business hours.

> [!NOTE]
> Support for this feature was introduced in [requirement set 1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13). See [clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Configure the manifest

To schedule the delivery of a message, your add-in must be able to activate in message compose mode. This is defined through the [MessageComposeCommandSurface](/javascript/api/manifest/extensionpoint#messagecomposecommandsurface) extension point in an add-in only manifest or the **mailCompose** "contexts" property in a [Unified manifest for Microsoft 365](../develop/unified-manifest-overview.md).

For further guidance on how to configure an Outlook add-in manifest, see [Office Add-in manifests](../develop/add-in-manifests.md).

## Access the delivery property of a message

The [item.delayDeliveryTime](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-delaydeliverytime-member) property returns a [DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime) object that provides you with methods to get or set the delivery date and time of a message.

## Get the delivery date and time of a message

To get the delivery date and time of a message in compose mode, call [item.delayDeliveryTime.getAsync](/javascript/api/outlook/office.delaydeliverytime#outlook-office-delaydeliverytime-getasync-member(1)) as shown in the following example. If a delivery date hasn't been set on a message yet, the call returns `0`. Otherwise, it returns a [JavaScript Date object](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date).

```javascript
// Gets the delivery date and time of a message.
Office.context.mailbox.item.delayDeliveryTime.getAsync((asyncResult) => {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log(asyncResult.error.message);
    return;
  }

  const deliveryDate = asyncResult.value;
  if (deliveryDate === 0) {
    console.log("Your message will be delivered immediately when you select Send.");
  } else {
    const date = new Date(deliveryDate);
    console.log(`Message delivery date and time: ${date.toString()}`);
  }
});
```

## Set the delivery date and time of a message

To delay the delivery of a message, pass a [JavaScript Date object](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) as a parameter to [item.delayDeliveryTime.setAsync](/javascript/api/outlook/office.delaydeliverytime#outlook-office-delaydeliverytime-setasync-member(1)) method, as shown in the following example.

```javascript
// Delays the delivery time by five minutes from the current time.
const currentTime = new Date().getTime();
const milliseconds = 5 * 60 * 1000;
const timeDelay = new Date(currentTime + milliseconds);
Office.context.mailbox.item.delayDeliveryTime.setAsync(timeDelay, (asyncResult) => {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log(asyncResult.error.message);
    return;
  }

  console.log("Message delivery has been scheduled.");
});
```

## Feature behavior and limitations

When you schedule the delivery of a message using the `item.delayDeliveryTime.setAsync` method, the delay is processed on the server. This allows the message to be sent even if the Outlook client isn’t running. However, because of this, the message doesn't appear in the Outbox folder, so you won't be able to edit the message or cancel its delivery after selecting **Send**. You'll be able to review the message from the **Sent Items** folder once the message is sent.

This behavior differs from a message scheduled using the native **Delay Delivery** option in the Outlook client, which processes the delay client-side. A message scheduled using this option appears in the **Outbox** folder and is only delivered if the Outlook client from which it was sent is running at the specified delivery time.

## Try sample snippets in Script Lab

Get the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) and try out the "Get and set message delivery (Message Compose)" sample snippet. To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

:::image type="content" source="../images/outlook-delay-delivery-script-lab.png" alt-text="The message delivery sample snippet in Script Lab.":::

## See also

- [Create Outlook add-ins for compose forms](compose-scenario.md)
