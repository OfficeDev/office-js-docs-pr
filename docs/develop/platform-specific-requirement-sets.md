---
title: Platform-specific requirement sets
description: Learn about platform-specific requirement sets.
ms.topic: how-to
ms.date: 11/18/2024
ms.localizationpriority: medium
---

# Platform-specific requirement sets

The Office Add-ins platform allows you to build solutions that extend Office applications and interact with content in Office documents. Your solution can run in Office across several platforms, including Windows, Mac, iPad, and in a browser. We've provided requirement sets that help you declare which APIs and platforms your add-in supports. Requirement sets are named groups of API members which are usually supported on all available platforms. However, with platform-specific requirement sets, APIs are provided sooner in the target platforms before support can be implemented in the remaining platforms.

Each application that supports Office Add-ins has its usual set of available platforms. For a comprehensive listing, see [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets). For the purpose of this discussion, we'll focus on Excel, Outlook, PowerPoint, and Word.

Excel, PowerPoint, and Word generally support Office Add-ins on the following platforms.

- Web browser
- Windows
- Mac
- iPad

Outlook support is usually available on the following:

- Web browser
- Windows
- Mac
- Android
- iOS

However, platform-specific requirement sets provide support for a subset of the usual platforms. For example, online-only requirement sets provide APIs that are only available when the add-in runs in a web browser. Similarly, desktop-only requirement sets provide APIs that may only be available when the add-in runs in Windows and Mac.

## Current platform-specific requirement sets

At present, platform-specific requirement sets are available in Excel and Word. Excel includes an online-only requirement set. Word includes online-only and desktop-only requirement sets. For the full list, see [Special requirement sets](/javascript/api/overview#special-requirement-sets).

Note that in Outlook, platform-specific behavior may be found in extension points. For example, [MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) and [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) are only available to add-ins running in Outlook on Android and iOS.

## Why use platform-specific requirement sets?

If your add-in has scenarios that target specific platforms, then a platform-specific requirement set may work for you. You don't need to wait for the APIs you care about to be implemented for the other platforms. Your add-in can use those APIs and you can ship your solution that much sooner to your customers. As you decide if to use platform-specific requirement sets, keep the following considerations in mind.

## TODO: more considerations

TODO

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office versions and requirement sets](office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md)
- [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets)
