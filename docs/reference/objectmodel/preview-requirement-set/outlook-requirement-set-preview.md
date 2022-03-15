---
title: Outlook add-in API preview requirement set
description: Features and APIs that are currently in preview for Outlook add-ins.
ms.date: 03/15/2022
ms.localizationpriority: medium
---

# Outlook add-in API preview requirement set

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center). "Configure preview access" is noted on this page for applicable features.
>
> For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview). "Request preview access" is noted on those features.

The preview requirement set includes all of the features of [requirement set 1.11](../requirement-set-1.11/outlook-requirement-set-1.11.md).

## Features in preview

The following features are in preview.

### Add-in activation on items protected by Information Rights Management (IRM)

Add-ins can now activate on IRM-protected items. To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office. See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.

**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)

<br>

---

---

### Additional calendar properties

#### [IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the all-day event property of an appointment in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Added a new object that represents the sensitivity of an appointment in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Added a new property that represents if an appointment is an all-day event.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Added a new property that represents the sensitivity of an appointment.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

<br>

---

---

### Delay delivery time

#### [Office.context.mailbox.item.delayDeliveryTime](office.context.mailbox.item.md#properties)

Added a new property that returns an object that allows you to manage the delivery date and time of a message in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-preview&preserve-view=true)

Added a new object that allows you to manage the delivery date and time of a message in Compose mode.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

<br>

---

---

### Event-based activation

This feature was released in [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). However, additional events are now available in preview. To learn more, refer to [Supported events](../../../outlook/autolaunch.md#supported-events).

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.AddinCommands.EventCompletedOptions.errorMessage](/javascript/api/office/office.addincommands.eventcompletedoptions?view=outlook-js-preview&preserve-view=true#office-office-addincommands-eventcompletedoptions-errormessage-member)

Added a new property to display an error message to the user if the handled event can't continue to execute. For an example, refer to the [Smart Alerts walkthrough](../../../outlook/smart-alerts-onmessagesend-walkthrough.md).

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

<br>

---

---

### Integration with actionable messages

#### [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)

<br>

---

---

### Office theme

#### [Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#office-office-context-officetheme-member)

Added ability to get Office theme.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

#### [Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

Added `OfficeThemeChanged` event to `Mailbox`.

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)

<br>

---

---

### Shared mailboxes

Feature support for shared folders (that is, delegate access) was released in [requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). However, support for shared mailboxes is now available in preview. To learn more, refer to [Enable shared folders and shared mailbox scenarios](../../../outlook/delegate-access.md).

**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
