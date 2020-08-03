---
title: Outlook add-in API Preview requirement set
description: 'Features and APIs that are currently in preview for Outlook add-ins.'
ms.date: 08/03/2020
localization_priority: Normal
---

# Outlook add-in API Preview requirement set

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center). "Configure preview access" is noted on this page for applicable features.
>
> For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview). "Request preview access" is noted on those features.

The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## Features in preview

The following features are in preview.

### Add-in activation on items protected by Information Rights Management (IRM)

Add-ins can now activate on IRM-protected items. To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office. See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.

**Available in**: Outlook on Windows, starting with build 13120.1000 (connected to a Microsoft 365 subscription)

<br>

---

---

### Additional calendar properties

#### [IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

Added a new object that represents the all-day event property of an appointment in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

Added a new object that represents the sensitivity of an appointment in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Added a new property that represents if an appointment is an all-day event.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Added a new property that represents the sensitivity of an appointment.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

<br>

---

---

### Append on send

To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).

#### [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### [ExtendedPermissions](../../manifest/extendedpermissions.md)

Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### Async versions of `display` APIs

#### [Office.context.mailbox.displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentformasync-itemid--options--callback-)

Added a new function to the `Mailbox` object that displays an existing appointment. This is the async version of the `displayAppointmentForm` method.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageformasync-itemid--options--callback-)

Added a new function to the `Mailbox` object that displays an existing message. This is the async version of the `displayMessageForm` method.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentformasync-parameters--options--callback-)

Added a new function to the `Mailbox` object that displays a new appointment form. This is the async version of the `displayNewAppointmentForm` method.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageformasync-parameters--options--callback-)

Added a new function to the `Mailbox` object that displays a new message form. This is the async version of the `displayNewMessageForm` method.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.item.displayReplyAllFormAsync](office.context.mailbox.item.md#methods)

Added a new function to the `Item` object that displays the "Reply all" form in Read mode. This is the async version of the `displayReplyAllForm` method.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.item.displayReplyFormAsync](office.context.mailbox.item.md#methods)

Added a new function to the `Item` object that displays the "Reply" form in Read mode. This is the async version of the `displayReplyForm` method.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

<br>

---

---

### Event-based activation

Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.

#### [LaunchEvent extension point](../../manifest/extensionpoint.md#launchevent-preview)

Added `LaunchEvent` extension point support to manifest. It configures event-based activation functionality.

**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))

#### [LaunchEvents manifest element](../../manifest/launchevents.md)

Added `LaunchEvents` element to manifest. It supports configuring event-based activation functionality.

**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))

#### [Runtimes manifest element](../../manifest/runtimes.md)

Added Outlook support to the `Runtimes` manifest element. It references the HTML and JavaScript files needed for event-based activation functionality.

**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))

<br>

---

---

### Get all custom properties

#### [CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

Added a new function to the `CustomProperties` object that gets all custom properties.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on Android, Outlook on iOS

<br>

---

---

### Integration with actionable messages

#### [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)

<br>

---

---

### Mail signature

#### [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### [Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Added a new function that disables the client signature for the sending mailbox in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

Added a new function that gets the compose type of a message in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### [Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Added a new function that checks if the client signature is enabled on the item in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### [Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

Added a new enum `ComposeType` available in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### Notification messages with actions

This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.

#### [Office.NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

Added a new property that enables you to add an `InsightMessage` notification with a custom action.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)

#### [Office.NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

Added a new object where you define a custom action for your `InsightMessage` notification.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)

#### [Office.MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype)

Added a new enum `ActionType`.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)

#### [Office.MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)

<br>

---

---

### Office theme

#### [Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Added ability to get Office theme.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Added `OfficeThemeChanged` event to `Mailbox`.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

<br>

---

---

### Single sign-on (SSO)

#### [OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
