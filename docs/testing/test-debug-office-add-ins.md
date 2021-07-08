---
title: Test and debug Office Add-ins
description: 'Learn how to test and debug your Office Add-in'
ms.date: 05/19/2021
localization_priority: Priority
---

# Test and debug Office Add-ins

This article contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.

## Test cross-platform and for multiple versions of Office

Office Add-ins run across major platforms, so you need to test an add-in in all the platforms where your users might be running Office. This usually includes Office on the web, Office on Windows (both subscription and one-time purchase), Office on Mac, Office on iOS, and (for Outlook add-ins) Office on Android. However, there may be some situations in which you can be sure that none of your users will be working on some platforms. For example, if you are making an add-in for a company that requires its users to work with Windows computers and subscription Office, then you don't need to test for Office on Mac or one-time purchase Windows.

> [!NOTE]
> On Windows computers, the version of Windows and Office will determine which browser control is used by add-ins. For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Add-ins marketed through AppSource go through a validation process that includes testing on all platforms. In addition, add-ins are tested for Office on the web with all major modern browsers, including Microsoft Edge (Chromium-based WebView2), Chrome, and Safari. Accordingly, you should test on these platforms and browsers before you submit to AppSource. For more information about validation, see [Commercial marketplace certification policies](/legal/marketplace/certification-policies), especially [section 1120.3](/legal/marketplace/certification-policies#11203-functionality), and the [Office Add-in application and availability page](../overview/office-add-in-availability.md).
>
> AppSource does not use Internet Explorer or the legacy version of Microsoft Edge (WebView1) to test add-ins in Office on the web. But if a significant number of your users will use these two browsers to open Office on the web, then you should test with them. For more information, see [Support Internet Explorer 11](../develop/support-ie-11.md) and [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues). Office still supports these browsers for add-ins, so if you think you've encountered a bug in how add-ins run in them, please create an issue for the [office-js](https://github.com/OfficeDev/office-js/issues/new/choose) repo.

## Sideload an Office Add-in for testing

You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product.

- [Sideload Office Add-ins on Windows](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Sideload Office Add-ins in Office on the web](sideload-office-add-ins-for-testing.md)

- [Sideload Office Add-ins on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md)

- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)

## Debug an Office Add-in

The procedure for debugging an Office Add-in varies by platform as well. Each of the following articles describes how to debug Office Add-ins on a specific platform.

- [Attach a debugger from the task pane (on Windows)](attach-debugger-from-task-pane.md)

- [Debug add-ins using F12 developer tools on Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md)

- [Debug Office Add-ins on iPad and Mac](debug-office-add-ins-on-ipad-and-mac.md)

- [Microsoft Office Add-in Debugger Extension for Visual Studio Code](debug-with-vs-extension.md)

## Validate an Office Add-in manifest

For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).

## Troubleshoot user errors

For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).
