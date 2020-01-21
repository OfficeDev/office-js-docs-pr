---
title: Outlook add-in API Preview requirement set
description: ''
ms.date: 01/31/2020
localization_priority: Priority
---

# Outlook add-in API Preview requirement set

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!IMPORTANT]
> This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## Features in preview

The following features are in preview.

### Integration with actionable messages

#### [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)

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

#### [MobileOnlineMeetingCommandSurface extension point](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

Added `MobileOnlineMeetingCommandSurface` extension point to manifest. It defines the online meeting integration.

**Available in**: Outlook on Android (connected to Office 365 subscription)

<br>

---

---

### SSO

#### [OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Added access to `getAccessToken`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)

## See also

- [Outlook add-ins](/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/outlook/add-ins/quick-start)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
