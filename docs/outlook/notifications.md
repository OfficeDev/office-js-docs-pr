---
title: Create notifications for your Outlook add-in
description: Learn about the types of notification messages you can create for your Outlook add-in.
ms.date: 05/29/2025
ms.localizationpriority: medium
---

# Create notifications for your Outlook add-in

Implement notification messages for your Outlook add-in to keep your users informed about important events, feedback, or errors with minimal disruption to their workflow.

> [!NOTE]
> Support for the notifications API was introduced in Mailbox requirement set 1.3. Additional features were introduced in later requirement sets. To determine if your client supports these requirement sets, see [Outlook client support](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support).

## Supported Outlook surfaces and modes

Notification messages are supported on messages and appointments in both read and compose modes. They're displayed above the body of the mail item.

:::image type="content" source="../images/outlook-notification.png" alt-text="An insight notification displayed in an appointment in compose mode.":::

To manage a notification on a mail item, call [Office.context.mailbox.item.notificationMessages](/javascript/api/requirement-sets/outlook/requirement-set-1.3/office.context.mailbox.item#properties) in your add-in's JavaScript code. This property returns a [NotificationMessages](/javascript/api/outlook/office.notificationmessages) object with methods to add, remove, get, or replace notifications. The following code shows how to use these methods to manage your add-in's notifications.

```javascript
const notificationMessages = Office.context.mailbox.item.notificationMessages;

// Sample informational message.
const notificationDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "This is a sample notification message.",
    icon: "icon-16",
    persistent: false
};

const notificationKey = "notification_01";

// Add a notification to the mail item.
notificationMessages.addAsync(notificationKey, notificationDetails, (result) => {
    console.log("Added an informational notification.");
});

// Get all the notifications of the mail item.
notificationMessages.getAllAsync((result) => {
    console.log(JSON.stringify(result.value));
});

// Replace a notification.
const newNotification = {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message: "This is a sample error message."
};

notificationMessages.replaceAsync(notificationKey, newNotification, (result) => {
    console.log("Replaced the existing notification.");
});

// Remove a notification.
notificationMessages.removeAsync(notificationKey, (result) => {
    console.log("Removed the notification.");
});
```

## Types of notifications

A notification consists of a unique identifier, an icon, and a message. Depending on the type, it could also include a **Dismiss** action or a custom action. There are different [types of notifications](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) you can display to the user to fit your particular scenario.

- [ErrorMessage](#errormessage)
- [InformationalMessage](#informationalmessage)
- [InsightMessage](#insightmessage)
- [ProgressIndicator](#progressindicator)

The following sections describe each notification type, including its [properties](/javascript/api/outlook/office.notificationmessagedetails) and supported platforms.

### ErrorMessage

:::row:::
    :::column:::
        **Description**
    :::column-end:::
    :::column span="3":::
        Alerts the user about an error or failed operation. For example, use the `ErrorMessage` type to notify the user that their personalized signature wasn't successfully added to a message.

:::image type="content" source="../images/outlook-error-notification.png" alt-text="An error message notification.":::
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Properties**
    :::column-end:::
    :::column span="3":::
        - Displays an error icon. This icon can't be customized.
        - Includes a **Dismiss** action to close the notification. If a user doesn't dismiss the error notification, it remains visible until the user sees it once before switching to another mail item.
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Minimum supported requirement set**
    :::column-end:::
    :::column span="3":::
        [1.3](/javascript/api/requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3)
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Supported platforms**
    :::column-end:::
    :::column span="3":::
        - Web
        - Windows (new and classic)
        - Mac
        - Android
        - iOS
    :::column-end:::
:::row-end:::

### InformationalMessage

:::row:::
    :::column:::
        **Description**
    :::column-end:::
    :::column span="3":::
        Provides information or feedback to the user. For example, use the `InformationalMessage` type to notify the user that their file upload completed successfully.

:::image type="content" source="../images/outlook-informational-notification.png" alt-text="An informational notification.":::
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Properties**
    :::column-end:::
    :::column span="3":::
        - Must specify an icon. Although an icon is required, the custom icon is currently displayed only in classic Outlook on Windows. On other platforms, an information icon is shown.
        - Includes a **Dismiss** action to close the notification.
        - Can be customized to persist even after a user switches to another mail item. The notification remains until the add-in removes it or the user selects **Dismiss**.
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Minimum supported requirement set**
    :::column-end:::
    :::column span="3":::
        [1.3](/javascript/api/requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3)
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Supported platforms**
    :::column-end:::
    :::column span="3":::
        - Web
        - Windows (new and classic)
        - Mac
        - Android
        - iOS
    :::column-end:::
:::row-end:::

### InsightMessage

:::row:::
    :::column:::
        **Description**
    :::column-end:::
    :::column span="3":::
         Provides information or feedback to the user with an option to perform an action. For example, use the `InsightMessage` type to recommend adding catering services to a meeting with external recipients.

:::image type="content" source="../images/outlook-insight-notification.png" alt-text="An insight message notification.":::
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Properties**
    :::column-end:::
    :::column span="3":::
        - Must specify an icon. Although an icon is required, the custom icon is displayed only in classic Outlook on Windows. On other platforms, an information icon is shown.
        - Includes an option to perform one [action](/javascript/api/outlook/office.notificationmessageaction). Currently, opening the add-in's task pane is the only supported action.
        - Includes a **Dismiss** action to close the notification.
        - Doesn't persist when a user switches to another mail item.
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Minimum supported requirement set**
    :::column-end:::
    :::column span="3":::
        [1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Supported platforms**
    :::column-end:::
    :::column span="3":::
        - Web
        - Windows (new and classic)
        - Mac
    :::column-end:::
:::row-end:::

### ProgressIndicator

:::row:::
    :::column:::
        **Description**
    :::column-end:::
    :::column span="3":::
        Indicates the progress of an add-in operation. For example, use the `ProgressIndicator` to inform the user that their file is in the process of being attached to the mail item.

:::image type="content" source="../images/outlook-progress-notification.png" alt-text="A progress indicator notification.":::
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Properties**
    :::column-end:::
    :::column span="3":::
        - In classic Outlook on Windows, displays a progress icon. On other platforms, displays an information icon. This icon can't be customized.
        - Doesn't persist when a user switches to another mail item.
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Minimum supported requirement set**
    :::column-end:::
    :::column span="3":::
        [1.3](/javascript/api/requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3)
    :::column-end:::
:::row-end:::
:::row:::
    :::column:::
        **Supported platforms**
    :::column-end:::
    :::column span="3":::
        - Web
        - Windows (new and classic)
        - Mac
        - Android
        - iOS
    :::column-end:::
:::row-end:::

## Feature behaviors

When creating and managing notifications for your add-in, be mindful of the following behaviors, limitations, and best practices.

### Maximum number of notifications per mail item

In Outlook on the web, on Windows (new and classic), and on Mac, you can add a maximum of five notifications per message. In Outlook on mobile devices, only one notification can be added to a message. Setting an additional notification replaces the existing one.

### InsightMessage limitations

Only one `InsightMessage` notification is allowed per add-in on a mail item. In Outlook on the web and new Outlook on Windows, the `InsightMessage` type is only supported in compose mode.

### Notification icons and unified manifest for Microsoft 365

If your add-in uses the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md), you can't customize the icon of an `InformationalMessage` or `InsightMessage` notification. The notification uses the first image specified in the ["icons"](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item#icons) array of the first [extensions.ribbons.tabs.groups.controls](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item) object of the manifest. Although this is the case, you must still specify a string in the [icon](/javascript/api/outlook/office.notificationmessagedetails#outlook-office-notificationmessagedetails-icon-member) property of your [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails) object (for example, "icon-16").

### Notification icons in Outlook on mobile devices

In compose mode, while the style of each notification type varies on other Outlook clients, notifications in Outlook on Android and on iOS all use the same style. The notification message always uses an information icon.

### Notifications for multiple selected messages

When managing notifications for multiple selected messages, only the `getAllAsync` method is supported. To learn more, see [Activate your Outlook add-in on multiple messages](item-multi-select.md).

### Best practices for ProgressIndicator notifications

When implementing a `ProgressIndicator` notification in your add-in, once the applicable operation or action completes, replace the progress notification with another notification type. This is a best practice to ensure that your users always get the latest status of an operation.

## Try the code example in Script Lab

Learn how you can use notifications in your add-in by trying out the [Work with notification messages](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/outlook/35-notifications/add-getall-remove.yaml) sample in [Script Lab for Outlook](https://appsource.microsoft.com/product/office/wa200001603). For more information on Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

## See also

- [Use the Office dialog API in Office Add-ins](../develop/dialog-api-in-office-add-ins.md)
- [Activate add-ins with events](../develop/event-based-activation.md)
