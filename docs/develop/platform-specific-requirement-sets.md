---
title: Understanding platform-specific requirement sets
description: Understand and learn how to use platform-specific requirement sets.
ms.topic: how-to
ms.date: 10/27/2025
ms.localizationpriority: medium
---

# Understanding platform-specific requirement sets

The Office Add-ins platform allows you to build solutions that extend Office applications and interact with content in Office documents. Your solution can run in Office across several platforms, including Windows, Mac, iPad, and in a browser. We've provided requirement sets that help you declare which APIs and platforms your add-in supports. Requirement sets are named groups of API members which are usually supported on all available platforms. However, with platform-specific requirement sets, APIs are implemented and made available first in the target platforms.

Each application that supports Office Add-ins has its usual set of available platforms. For a comprehensive listing, see [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets). For the purpose of this discussion, we'll focus on Excel, Outlook, PowerPoint, and Word.

Cross-platform requirement sets are available on Windows, Mac, and in a browser. Depending on the features being made available, a requirement set may also be supported on iPad or mobile platforms.

However, platform-specific requirement sets provide support for a subset of the usual platforms. For example, online-only requirement sets provide APIs that are only available when the add-in runs in a web browser. Similarly, desktop-only requirement sets provide APIs that may only be available when the add-in runs in Windows and Mac. See the specific requirement set page for actual platform support.

## Current platform-specific requirement sets

At present, platform-specific requirement sets are available in Excel and Word. Excel and Word include both online-only and desktop-only requirement sets. For the full list, see [Special requirement sets](/javascript/api/overview#special-requirement-sets).

Note that in Outlook, platform-specific behavior may be found in extension points. For example, [MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) and [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) are only available to add-ins running in Outlook on Android and on iOS.

## Why platform-specific requirement sets?

We're providing platform-specific requirement sets for a few reasons.

1. **Feature availability.** Some features aren't implemented in the Office applications UI on a particular platform. As such, the API can only be made available on supported platforms. Having these types of APIs in a platform-specific requirement set means that developers can use those APIs in their add-ins. This is especially useful for cases where the feature may never be implemented in other platforms.
1. **Platform-specific add-ins.** Developers who have add-ins focused on a particular platform don't need to wait for those APIs to be implemented in other platforms. These developers are able to incorporate those APIs into their solutions and ship to their customers much sooner.
1. **Tailored experiences.** Customers can use an Office application differently depending on the platform for several reasons, like feature availability or comfort level, for example. Let's say that on the Windows version, a customer completes one set of tasks but on an iPad, they complete a different set of tasks. You can have your add-in provide a tailored experience based on your users' usual scenarios per platform.

To help you decide if platform-specific requirement sets can work for you, consider the following.

## API promotion to cross-platform requirement set

When APIs in a platform-specific requirement set are supported cross-platform, they're added to the next requirement set targeted for release. Even after the new requirement is made generally available, those APIs *still remain* in the platform-specific requirement set.

## How to use a platform-specific requirement set

The following sections describe where you can specify your minimum requirement set. For more information about these options, see [Specify which Office versions and platforms can host your add-in](specify-office-hosts-and-api-requirements.md).

### Manifest

When you note a requirement set in the [Set element](/javascript/api/manifest/set) of your add-in manifest, you're indicating the minimum set of APIs that your add-in needs. Combined with supported Office host applications and other information, this determines whether or not your add-in activates in an Office client.

When you declare a platform-specific requirement set, your add-in activates only when it's run in Office on that platform. For example, if you have the WordApiDesktop 1.1 requirement set in your manifest, your add-in will only activate in Word on Windows and on Mac.

Keep in mind that in the case where the APIs become supported cross-platform, you'll need to update your add-in manifest to add a cross-platform requirement set and remove the platform-specific requirement set. If your add-in is available in Microsoft Marketplace, you'll need to resubmit it for validation.

### Code

Another option is to implement a runtime check in your code. This way, you can make new features available to your customers on those platforms. The runtime check also ensures that the platform-specific code doesn't run on unsupported platforms and cross-platform features continue to work for your customers. The following code is an example that checks for a specific requirement set.

```javascript
if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.1")) {
   // Any API exclusive to this WordApiDesktop requirement set.
}
```

Whenever platform-specific APIs become available cross-platform, enable your customers on all supported platforms to use those features by implementing one of the following options.

- Remove the runtime check. But note that customers on older Office clients, especially on Windows, may hit errors if their client doesn't support the new APIs yet.
- Update the runtime code to check for the cross-platform requirement set.

A variation is to do a runtime check for a particular API. This means that the encapsulated code should run on any platforms that support that API. If the API was first released in a platform-specific requirement set then promoted to a cross-platform one, you shouldn't need to update your code unless you made assumptions about the supported platforms. The following code is an example.

```javascript
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```

## Notify customers on Microsoft Marketplace

If your add-in is in Microsoft Marketplace or the Office store, be sure to notify customers about any platform-specific behavior.

**Details + support** > **Products supported** on your add-in's Microsoft Marketplace page should automatically show the appropriate supported platforms based on the requirements you declared in the manifest.

However, if your add-in is supported cross-platform but you also implemented platform-specific behaviors, you should point out those feature differences in the **Overview** section on your add-in's Microsoft Marketplace page.

## Exceptions

The following are exceptions to the approach described.

### Online-only requirement sets

An online-only requirement set is a superset of the latest numbered requirement set. For each Office application with an online-only requirement set, `1.1` is the only version. It's invalid to specify an online-only requirement set in the [Set element](/javascript/api/manifest/set) of your add-in manifest.

To check for APIs that are only supported in these requirement sets and to prevent your add-in from trying to run the code on unsupported platforms, add code similar to the following:

```javascript
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

```javascript
if (Office.context.requirements.isSetSupported("WordApiOnline", "1.1")) {
   // Any API exclusive to the WordApiOnline requirement set.
}
```

When APIs in an online-only requirement set are supported cross-platform, they're added to the next released requirement set. After the new requirement set is made generally available, those APIs are *removed* from the online-only requirement set.

Follow the guidance in the earlier [Code](#code) section to adjust your add-in implementation accordingly.

### Desktop-only HiddenDocument requirement sets in Word

It's important to note that while the HiddenDocument requirement sets in Word are desktop-only, it's invalid to specify a HiddenDocument requirement set in the [Set element](/javascript/api/manifest/set) of your add-in manifest.

To check for APIs that are only supported in these requirement sets and to prevent your add-in from trying to run the code on unsupported platforms, add code similar to the following:

```javascript
if (Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.5")) {
   // Any API exclusive to this WordApiHiddenDocument requirement set.
}
```

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office versions and requirement sets](office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md)
- [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets)
