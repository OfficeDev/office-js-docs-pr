---
title: Outlook add-in API Preview requirement set
description: 'Features and APIs that are currently in preview for Outlook add-ins.'
ms.date: 05/13/2020
localization_priority: Normal
---

# Outlook add-in API Preview requirement set

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> You may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview). "Request access" is noted on this page for applicable features.
>
> For other features, you may be able to preview them in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center). "Configure preview access" is noted on those features.

The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## Features in preview

The following features are in preview.

### Additional calendar properties

#### [IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

Added a new object that represents the all-day event property of an appointment in Compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

#### [Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

Added a new object that represents the sensitivity of an appointment in Compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

#### [Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Added a new property that represents if an appointment is an all-day event.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

#### [Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Added a new property that represents the sensitivity of an appointment.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

#### [Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

<br>

---

---

### Append on send

#### [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### [ExtendedPermissions](../../manifest/extendedpermissions.md)

Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### Integration with actionable messages

#### [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)

<br>

---

---

### Mail signature

#### [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

#### [Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Added a new function that disables the client signature for the sending mailbox in Compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

#### [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

Added a new function that gets the compose type of a message in Compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

#### [Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Added a new function that checks if the client signature is enabled on the item in Compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

#### [Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

Added a new enum `ComposeType` available in Compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Request access](https://aka.ms/OWAPreview))

<br>

---

---

### Office theme

#### [Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Added ability to get Office theme.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Added `OfficeThemeChanged` event to `Mailbox`.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

<br>

---

---

### Online meeting provider integration

Added support for online-meeting integration in Outlook mobile add-ins. See [Create an Outlook mobile add-in for an online-meeting provider](../../../outlook/online-meeting.md) to learn more.

#### [MobileOnlineMeetingCommandSurface extension point](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

Added `MobileOnlineMeetingCommandSurface` extension point to manifest. It defines the online meeting integration.

**Available in**: Outlook on Android (connected to Office 365 subscription)

<br>

---

---

### SSO

#### [OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
