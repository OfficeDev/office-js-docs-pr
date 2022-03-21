---
title: Outlook add-in API requirement set 1.10
description: 'Requirement set 1.10 for Outlook add-in API.'
ms.date: 11/04/2021
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.10

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.10?

Requirement set 1.10 includes all of the features of [requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md). It added the following features.

- Added new APIs for [event-based activation](../../../outlook/autolaunch.md) and mail signature features.
- Added support for the [OfficeRuntime.Storage](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true) object with the event-based activation feature.
- Added ability to include a custom action on a notification message.

### Change log

- Added [LaunchEvent extension point](/javascript/api/manifest/extensionpoint#launchevent): Adds a new supported type of ExtensionPoint. It configures event-based activation functionality.
- Added [LaunchEvents manifest element](/javascript/api/manifest/launchevents): Adds a manifest element to support configuring event-based activation functionality.
- Modified [Runtimes manifest element](/javascript/api/manifest/runtimes): Adds Outlook support. It references the HTML and JavaScript files needed for event-based activation functionality.
- Added [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-1.10&preserve-view=true#outlook-office-body-setsignatureasync-member(1)): Adds a new function to the `Body` object. It adds or replaces the signature in the item body in Compose mode.
- Added [Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods): Adds a new function that disables the client signature for the sending mailbox in Compose mode.
- Added [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#outlook-office-messagecompose-getcomposetypeasync-member(1)): Adds a new function that gets the compose type of a message in Compose mode.
- Added [Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods): Adds a new function that checks if the client signature is enabled on the item in Compose mode.
- Added [Office.MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype?view=outlook-js-1.10&preserve-view=true): Adds a new enum. It represents the type of custom action in a notification message.
- Added [Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-1.10&preserve-view=true): Adds a new enum available in Compose mode.
- Added [Office.MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.10&preserve-view=true): Adds a new type to the `ItemNotificationMessageType` enum. It represents a notification message with a custom action.
- Added [Office.NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction?view=outlook-js-1.10&preserve-view=true): Adds a new object so you can define a custom action for your `InsightMessage` notification.
- Added [Office.NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.10&preserve-view=true#outlook-office-notificationmessagedetails-actions-member): Adds a new property that enables you to add an `InsightMessage` notification with a custom action.
- Modified [OfficeRuntime.Storage](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true): Adds Outlook support but only with the event-based activation feature.

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
