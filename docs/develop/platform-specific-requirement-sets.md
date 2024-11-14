---
title: Platform-specific requirement sets
description: Learn about platform-specific requirement sets.
ms.topic: how-to
ms.date: 11/18/2024
ms.localizationpriority: medium
---

# Platform-specific requirement sets

The Office Add-ins platform allows you to build solutions that extend Office applications and interact with content in Office documents. Your solution can run in Office across several platforms, including Windows, Mac, iPad, and in a browser. We've provided requirement sets that help you declare which APIs and platforms your add-in supports. Requirement sets are named groups of API members which are usually supported on all available platforms. However, with platform-specific requirement sets, APIs are implemented and made available first in the target platforms.

Each application that supports Office Add-ins has its usual set of available platforms. For a comprehensive listing, see [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets). For the purpose of this discussion, we'll focus on Excel, Outlook, PowerPoint, and Word.

Cross-platform requirement sets are at least available on Windows, Mac, and in a browser. Depending on the features being made available, a requirement set may also be supported on iPad or mobile platforms.

However, platform-specific requirement sets provide support for a subset of the usual platforms. For example, online-only requirement sets provide APIs that are only available when the add-in runs in a web browser. Similarly, desktop-only requirement sets provide APIs that may only be available when the add-in runs in Windows and Mac.

## Current platform-specific requirement sets

At present, platform-specific requirement sets are available in Excel and Word. Excel includes an online-only requirement set. Word includes online-only and desktop-only requirement sets. For the full list, see [Special requirement sets](/javascript/api/overview#special-requirement-sets).

Note that in Outlook, platform-specific behavior may be found in extension points. For example, [MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) and [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) are only available to add-ins running in Outlook on Android and on iOS.

## Why platform-specific requirement sets?

We're providing platform-specific requirement sets for a few reasons.

1. **Feature availability.** Some features aren't implemented in the Office applications UI on a particular platform. As such, the API can only be made available on supported platforms. Having these types of APIs to a platform-specific requirement set means that developers can use those APIs in their add-ins. This is especially useful for cases where the feature may never be implemented in other platforms.
1. **Platform-specific add-ins.** Developers who have add-ins focused on a particular platform don't need to wait for those APIs to be implemented in other platforms. These developers are able to incorporate those APIs into their solutions and ship to their customers much sooner.
1. **Tailored experiences.** Customers can use an Office application differently depending on the platform for several reasons, like feature availability or comfort level, for example. Let's say that on the Windows version, a customer completes one set of tasks but on an iPad, they complete a different set of tasks. You can have your add-in provide a tailored experience based on your users' usual scenarios per platform.

To help you decide if platform-specific requirement sets can work for you, the following are some considerations.

## TODO: Some considerations

TODO

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office versions and requirement sets](office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md)
- [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets)
