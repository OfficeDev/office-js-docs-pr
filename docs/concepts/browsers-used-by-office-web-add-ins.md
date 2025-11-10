---
title: Browsers and webview controls used by Office Add-ins
description: Specifies how the operating system and Office version determine what webview is used by Office Add-ins.
ms.topic: concept-article
ms.date: 11/06/2025
ms.localizationpriority: medium
---

# Browsers and webview controls used by Office Add-ins

Office Add-ins are web applications that are displayed using iframes when running in Office on the web. In Office for desktop and mobile clients, Office Add-ins use an embedded browser control (also known as a webview). Add-ins also need a JavaScript engine to run the JavaScript. Both the embedded browser and the engine are supplied by a browser installed on the user's computer. In this article, "webview" refers to the combination of a webview control and a JavaScript engine. Which webview is used depends on the computer's operating system.

## Browsers by platform

The following table specifies which browser is used for the various platforms and operating systems.

|OS|Office version|Browser|
|:-----|:-----|:-----|
|any|Office on the web|The browser in which Office is opened.|
|Windows|any|Microsoft Edge (Chromium-based) with WebView2*|
|Mac|any|Safari with WKWebView|
|iOS|any|Safari with WKWebView|
|Android|any|Chrome|

\* WebView2 is installed with Office by default for [supported versions of Office](https://support.microsoft.com/office/818c68bc-d5e5-47e5-b52f-ddf636cf8e16).

> [!IMPORTANT]
> [Conditional Access](/azure/active-directory/conditional-access/overview) isn't supported for Office Add-ins on iOS or Android. Those add-ins use the Safari-based WKWebView or the Android-based WebView, not an Edge-based browser control.

## Troubleshoot WebView2 issues

### Scroll bar doesn't appear in task pane

By default, scroll bars in WebView2 are hidden until hovered over. To ensure that the scroll bar is always visible, the CSS styling that applies to the `<body>` element of the pages in the task pane should include the [-ms-overflow-style](https://devdoc.net/web/developer.mozilla.org/en-US/docs/Web/CSS/-ms-overflow-style.html) property and it should be set to `scrollbar`.

### Get errors trying to download a PDF file

Directly downloading blobs as PDF files in an add-in isn't supported with WebView2. The workaround is to create a simple web application that downloads blobs as PDF files. In your add-in, call the `Office.context.ui.openBrowserWindow(url)` method and pass the URL of the web application. This will open the web application in a browser window outside of Office.

## WIP-protected documents

There's an extra step needed for Add-ins to run in a document with [WIP (Windows Information Protection)](/windows/uwp/enterprise/wip-hub) and use **WebView2 (Microsoft Edge Chromium-based)**. Add the WebView2 process, **msedgewebview2.exe**, to the protected apps list in your enterprise's WIP policy. An admin [adds this WIP policy through Intune](/mem/intune/apps/windows-information-protection-policy-create#to-add-a-wip-policy) with the following values.

- **Name**: Webview2
- **Publisher**: O=MICROSOFT CORPORATION, L=REDMOND, S=WASHINGTON, C=US
- **Product Name**: MICROSOFT EDGE WEBVIEW2
- **File**: MSEDGEWEBVIEW2.EXE
- **Min Version**: *
- **Max Version**: *

To determine if a document is WIP-protected, follow these steps.

1. Open the file.
1. Select the **File** tab on the ribbon.
1. Select **Info**.
1. In the upper section of the **Info** page, just below the file name, a WIP-enabled document will have a briefcase icon followed by **Managed by Work (...)**.

## See also

- [Requirements for Running Office Add-ins](requirements-for-running-office-add-ins.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
