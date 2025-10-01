---
title: Manage the delivery date and time of a message
description: Learn how to schedule message delivery with the Office JavaScript API.
ms.date: 10/02/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Manage the delivery date and time of a message

Learn to build an Outlook add-in that schedules and sends messages even when the Outlook client is closed or the device is turned off. With the Office JavaScript API, your users can conveniently schedule email marketing campaigns or time messages for delivery during colleagues' or customers' business hours.

> [!NOTE]
> Support for this feature was introduced in [requirement set 1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13). See [clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Try it out

See the message delivery API in action. Install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) and try out the "Get and set message delivery (Message Compose)" sample snippet. To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

:::image type="content" source="../images/outlook-delay-delivery-script-lab.png" alt-text="The message delivery sample snippet in Script Lab.":::

## Configure the manifest

To schedule the delivery of a message, your add-in must be able to activate in message compose mode. This is defined through the [MessageComposeCommandSurface](/javascript/api/manifest/extensionpoint#messagecomposecommandsurface) extension point in an add-in only manifest or the **mailCompose** `"contexts"` property in a [Unified manifest for Microsoft 365](../develop/unified-manifest-overview.md).

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

When you schedule the delivery of a message using the `item.delayDeliveryTime.setAsync` method, the delay is processed on the server. This allows the message to be sent even if the Outlook client isn’t running. In classic Outlook on Windows, the message doesn't appear in the **Outbox** folder, so you won't be able to edit the message or cancel its delivery after selecting **Send**. However, you'll be able to review the message from the **Sent Items** folder. In Outlook on the web, on Mac, and in [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), the message appears in the **Drafts** folder until the scheduled delivery time. While it's in the **Drafts** folder, you'll be able to edit the message before it's sent.

The `item.delayDeliveryTime.setAsync` behavior differs from a message scheduled using the native **Delay Delivery** option in the Outlook client, which processes the delay client-side. A message scheduled using this option appears in the **Outbox** folder and is only delivered if the Outlook client from which it was sent is running at the specified delivery time.

## See also

- [Create Outlook add-ins for compose forms](compose-scenario.md)
