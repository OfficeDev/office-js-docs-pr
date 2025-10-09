---
title: Test Office Add-ins
description: Learn how to test your Office Add-in.
ms.date: 07/12/2024
ms.localizationpriority: high
---

# Test Office Add-ins

This article contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.

## Test cross-platform and for multiple versions of Office

Office Add-ins run across major platforms, so you need to test an add-in in all the platforms where your users might be running Office. This usually includes Office on the web, Office on Windows (both perpetual and Microsoft 365 subscription), Office on Mac, Office on iOS, and (for Outlook add-ins) Office on Android. However, there may be some situations in which you can be sure that none of your users will be working on some platforms. For example, if you're making an add-in for a company that requires its users to work with Windows computers and subscription Office, then you don't need to test for Office on Mac or perpetual Office on Windows.

> [!NOTE]
> On Windows computers, the version of Windows and Office will determine which browser or webview control is used by add-ins. For more information, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). For brevity hereafter, this article uses "browser control" to mean "browser or webview control".

### Add-ins tested for Office on the web

Add-ins are tested for Office on the web with all major modern browsers, including Microsoft Edge (Chromium-based WebView2), Chrome, and Safari. Accordingly, you should test on these platforms and browsers before you submit to [Microsoft Marketplace](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center). For more information about validation, see [Commercial marketplace certification policies](/legal/marketplace/certification-policies), especially [section 1120.3](/legal/marketplace/certification-policies#11203-functionality), and the [Office Add-in application and availability page](/javascript/api/requirement-sets).

Office on the web no longer opens in Internet Explorer or Microsoft Edge Legacy (EdgeHTML). Consequently, Microsoft Marketplace doesn't test Office on the web on these browsers. Office still supports these browsers for add-in runtimes, so if you think you've encountered a bug in how add-ins run in them, please create an issue in the [office-js](https://github.com/OfficeDev/office-js/issues) repository. For more information, see [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md) and [Troubleshoot EdgeHTML and WebView2 (Microsoft Edge) issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshoot-edgehtml-and-webview2-microsoft-edge-issues).

### Add-ins tested for Office on Windows

Some Office versions on Windows still use the webview controls that come with Internet Explorer and Microsoft Edge Legacy. Microsoft Marketplace tests whether your add-in supports these browser controls. If your add-in doesn't support these browser controls, Microsoft Marketplace only issues a warning and doesn't reject your add-in. In this instance, we recommend configuring a graceful failure message on your add-in for a smoother user experience. For further guidance, see [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md).

## Sideload an Office Add-in for testing

You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product.

> [!NOTE]
> Office Add-ins that use the unified manifest for Microsoft 365 are *directly* supported in Office on the web, in [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), and in Office on Windows connected to a Microsoft 365 subscription, Version 2304 (Build 16320.00000) or later. When the app package that contains the unified manifest is sideloaded to a platform that doesn't directly support that type of manifest then an add-in only manifest is generated from the unified manifest and this manifest is the one that's sideloaded.  

- [Sideload Office Add-ins in Office on the web](sideload-office-add-ins-for-testing.md)
- [Sideload Office Add-ins on Windows](test-debug-non-local-server.md)
- [Sideload Office Add-ins on Mac](sideload-an-office-add-in-on-mac.md)
- [Sideload Office Add-ins on iPad](sideload-an-office-add-in-on-ipad.md)
- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)

## Unit testing

For information about how to add unit tests to your add-in project, see [Unit testing in Office Add-ins](unit-testing.md).

## Debug an Office Add-in

The procedure for debugging an Office Add-in varies based on your platform and environment. For more information, see [Debug Office Add-ins](debug-add-ins-overview.md).

## Validate an Office Add-in manifest

For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).

## Troubleshoot user errors

For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).
