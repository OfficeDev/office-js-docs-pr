---
title: Browsers and webview controls used by Office Add-ins
description: Specifies how the operating system and Office version determine what webview is used by Office Add-ins.
ms.topic: concept-article
ms.date: 10/17/2024
ms.localizationpriority: medium
---

# Browsers and webview controls used by Office Add-ins

Office Add-ins are web applications that are displayed using iframes when running in Office on the web. In Office for desktop and mobile clients, Office Add-ins use an embedded browser control (also known as a webview). Add-ins also need a JavaScript engine to run the JavaScript. Both the embedded browser and the engine are supplied by a browser installed on the user's computer. In this article, "webview" refers to the combination of a webview control and a JavaScript engine.

Which webview is used depends on:

- The computer's operating system.
- Whether the add-in is running in Office on the web, in Office downloaded from a Microsoft 365 subscription, or in perpetual Office 2016 or later.
- Within the perpetual versions of Office on Windows, whether the add-in is running in the "retail" or "volume-licensed" variation.

> [!IMPORTANT]
> **Webviews from Internet Explorer and Microsoft Edge Legacy are still used in Office Add-ins**
>
> Some combinations of platforms and Office versions, including volume-licensed perpetual versions through Office 2019, still use the webview controls that come with Internet Explorer 11 (called "Trident") and Microsoft Edge Legacy (called "EdgeHTML") to host add-ins, as explained in this article. Internet Explorer 11 was disabled in Windows 10 and Windows 11 in February 2023, and the UI for launching it was removed; but it's still installed on those operating systems. So, Trident and other functionality from Internet Explorer can still be called programmatically by Office.
>
> We recommend (but don't require) that you continue to support these combinations, at least in a minimal way, by providing users of your add-in a graceful failure message when your add-in is launched in one of these webviews. Keep these additional points in mind:
>
> - Office on the web no longer opens in Internet Explorer or Microsoft Edge Legacy. Consequently, [AppSource](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center) doesn't test add-ins in Office on these web browsers.
> - AppSource still tests for combinations of platform and Office *desktop* versions that use Trident or EdgeHTML. However, it only issues a warning when the add-in doesn't support these webviews; the add-in isn't rejected by AppSource.
> - The [Script Lab tool](../overview/explore-with-script-lab.md) no longer supports Trident.
>
> For more information about supporting Trident or EdgeHTML, including configuring a graceful failure message on your add-in, see [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md).

The following sections specify which browser is used for the various platforms and operating systems.

## Non-Windows platforms

For these platforms, the platform alone determines the browser that's used.

|OS|Office version|Browser|
|:-----|:-----|:-----|
|any|Office on the web|The browser in which Office is opened.<br>(But note that Office on the web will not open in Internet Explorer.<br>Attempting to do so opens Office on the web in Edge.) |
|Mac|any|Safari with WKWebView|
|iOS|any|Safari with WKWebView|
|Android|any|Chrome|

> [!IMPORTANT]
> [Conditional Access](/azure/active-directory/conditional-access/overview) isn't supported for Office Add-ins on iOS or Android. Those add-ins use the Safari-based WKWebView or the Android-based WebView, not an Edge-based browser control.

## Windows

An add-in running on Windows might use any of three different webviews:

- **WebView2**, which is provided by Microsoft Edge (Chromium-based).
- **EdgeHTML**, which is provided by Microsoft Edge Legacy.
- **Trident+**, which is provided by Internet Explorer 11. The "+" on the end indicates that Office Add-ins use additional functionality from Internet Explorer 11 that isn't built into Trident itself.

### Perpetual versions of Office on Windows

For perpetual versions of Office on Windows, the browser that's used is determined by the Office version, whether the license is retail or volume-licensed, and whether the Edge WebView2 (Chromium-based) is installed. The version of Windows doesn't matter, but note that Office Add-ins aren't supported on versions earlier than Windows 7 and Office 2021 and later aren't supported on versions earlier than Windows 10.

To determine whether Office 2016 or Office 2019 is retail or volume-licensed, use the format of the Office version and build number. (For Office 2021 and later, the distinction between volume-licensed and retail doesn't matter.)

- **Retail**: For both Office 2016 and 2019, the format is `YYMM (xxxxx.xxxxxx)`, ending with two blocks of five digits; for example, `2206 (Build 15330.20264)`.
- **Volume-licensed**:
  - For Office 2016, the format is `16.0.xxxx.xxxxx`, ending with two blocks of *four* digits; for example, `16.0.5197.1000`.
  - For Office 2019, the format is `1808 (xxxxx.xxxxxx)`, ending with two blocks of *five* digits; for example, `1808 (Build 10388.20027)`. Note that the year and month is always `1808`.

| Office version | Retail vs. Volume-licensed | WebView2 installed? | Browser |
|:-----|:-----|:-----|:-----|
| Office 2024 | Doesn't matter | Yes<sup>1</sup> | WebView2 (Microsoft Edge<sup>2</sup> Chromium-based) |
| Office 2021 | Doesn't matter | Yes<sup>1</sup> | WebView2 (Microsoft Edge<sup>2</sup> Chromium-based) |
| Office 2019 | Retail | Yes<sup>1</sup> | WebView2 (Microsoft Edge<sup>2</sup> Chromium-based) |
| Office 2019 | Retail | No | EdgeHTML (Microsoft Edge Legacy)<sup>2, 3</sup></br>If Edge isn't installed, Trident+ (Internet Explorer 11) is used. |
| Office 2019 | Volume-licensed | Doesn't matter | Trident+ (Internet Explorer 11) |
| Office 2016 | Retail | Yes<sup>1</sup> | WebView2 (Microsoft Edge<sup>2</sup> Chromium-based) |
| Office 2016 | Retail | No | EdgeHTML (Microsoft Edge Legacy)<sup>2, 3</sup></br>If Edge isn't installed, Trident+ (Internet Explorer 11) is used. |
| Office 2016 | Volume-licensed | Doesn't matter | Trident+ (Internet Explorer 11) |

<sup>1</sup> On Windows versions prior to Windows 11, the WebView2 control must be installed so that Office can embed it. It's installed with perpetual Office 2021 or later; but it isn't automatically installed with Microsoft Edge. If you have an earlier version of perpetual Office, use the instructions for installing the control at [Microsoft Edge WebView2 / Embed web content ... with Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

<sup>2</sup> When you use either EdgeHTML or WebView2, the Windows Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane. In Trident+, the Narrator reads the title bar of the task pane, which comes from the add-in name that's specified in the add-in's manifest.

<sup>3</sup> If your add-in uses an add-in only manifest and includes the **\<Runtimes\>** element in the manifest or it uses the unified manifest and it includes an "extensions.runtimes.lifetime" property, then it won't use EdgeHTML. If the conditions for using WebView2 are met, then the add-in uses WebView2. Otherwise, it uses Trident+. For more information, see [Runtimes](/javascript/api/manifest/runtimes) and [Activate add-ins with events](../develop/event-based-activation.md).

### Microsoft 365 subscription versions of Office on Windows

For subscription Office on Windows, the browser that's used is determined by the operating system, the Office version, and whether the WebView2 control is installed.

|OS|Office version| WebView2 installed?|Browser|
|:-----|:-----|:-----|:-----|
|<ul><li>Windows 11</li><li>Windows 10</li><li>Windows 8.1</li><li>Windows Server 2022</li><li>Windows Server 2019</li><li>Windows Server 2016</li></ul>| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>1</sup>| Yes<sup>2</sup>|  WebView2 (Microsoft Edge<sup>3</sup> Chromium-based) |
|<ul><li>Window 11</li><li>Windows 10 ver.&nbsp;>=&nbsp;1903</li></ul>| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>1</sup>| No |EdgeHTML (Microsoft Edge Legacy)<sup>3, 4</sup>|
|<ul><li>Windows 11</li><li>Windows 10 ver.&nbsp;>=&nbsp;1903</li></ul>| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.11629&nbsp;*AND*&nbsp;<&nbsp;16.0.13530.20424<sup>1</sup>| Doesn't matter|EdgeHTML (Microsoft Edge Legacy)<sup>3, 4</sup>|
|<ul><li>Windows 11</li><li>Windows 10 ver.&nbsp;>=&nbsp;1903</li></ul>| Microsoft 365 ver.&nbsp;<&nbsp;16.0.11629<sup>1</sup>| Doesn't matter|Trident+ (Internet Explorer 11)|
|<ul><li>Windows 10 ver.&nbsp;<&nbsp;1903</li><li>Windows 8.1</li></ul>| Microsoft 365 | No| Trident+ (Internet Explorer 11)|
|<ul><li>Windows 7</li></ul>| Microsoft 365| Doesn't matter | Trident+ (Internet Explorer 11)|

<sup>1</sup> See the [update history page](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) for more details.

<sup>2</sup> On Windows versions prior to Windows 11, the WebView2 control must be installed so that Office can embed it. It's installed with Microsoft 365, Version 2101 or later, but it isn't automatically installed with Microsoft Edge. If you have an earlier version of Microsoft 365, use the instructions for installing the control at [Microsoft Edge WebView2 / Embed web content ... with Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/). On Microsoft 365 builds before 16.0.14326.xxxxx, you must also create the registry key **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** and set its value to `dword:00000001`.

<sup>3</sup> When you use either EdgeHTML or WebView2, the Windows Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane. In Trident+, the Narrator reads the title bar of the task pane, which comes from the add-in name that's specified in the add-in's manifest.

<sup>4</sup> If your add-in uses an add-in only manifest and includes the **\<Runtimes\>** element in the manifest or it uses the unified manifest and it includes an "extensions.runtimes.lifetime" property, then it won't use EdgeHTML. If the conditions for using WebView2 are met, then the add-in uses WebView2. Otherwise, it uses Trident+. For more information, see [Runtimes](/javascript/api/manifest/runtimes) and [Activate add-ins with events](../develop/event-based-activation.md).

## Working with Trident+ (Internet Explorer 11)

Trident+ doesn't support JavaScript versions later than ES5. If any of your add-in's users have platforms that use Trident+, then to use the syntax and features of ECMAScript 2015 or later, you have two options.

- Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).
- Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.

For more information about these options, see [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md).

Also, Trident+ doesn't support some HTML5 features such as media, recording, and location. To learn more, see [Determine the webview the add-in is running in at runtime](../develop/support-ie-11.md#determine-the-webview-the-add-in-is-running-in-at-runtime).

## Troubleshoot EdgeHTML and WebView2 (Microsoft Edge) issues

### Service Workers aren't working

Office Add-ins don't support Service Workers when EdgeHTML is used. They're supported with WebView2.

### Scroll bar doesn't appear in task pane

By default, scroll bars in EdgeHTML and WebView2 are hidden until hovered over. To ensure that the scroll bar is always visible, the CSS styling that applies to the `<body>` element of the pages in the task pane should include the [-ms-overflow-style](https://devdoc.net/web/developer.mozilla.org/en-US/docs/Web/CSS/-ms-overflow-style.html) property and it should be set to `scrollbar`.

### When debugging with the Microsoft Edge DevTools, the add-in crashes or reloads

Setting breakpoints in the [Microsoft Edge DevTools](https://apps.microsoft.com/detail/9mzbfrmz0mnj) for EdgeHTML can cause Office to think that the add-in is hung. It will automatically reload the add-in when this happens. To prevent this, add the following Registry key and value to the development computer: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### When the add-in tries to open, get "ADD-IN ERROR We can't open this add-in from the localhost" error

One known cause is that EdgeHTML requires that localhost be given a loopback exemption on the development computer. Follow the instructions at [Cannot open add-in from localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### Get errors trying to download a PDF file

Directly downloading blobs as PDF files in an add-in isn't supported with EdgeHTML or WebView2. The workaround is to create a simple web application that downloads blobs as PDF files. In your add-in, call the `Office.context.ui.openBrowserWindow(url)` method and pass the URL of the web application. This will open the web application in a browser window outside of Office.

## WIP-protected documents

There's an extra step needed for Add-ins to run in a document with [WIP (Windows Information Protection)](/windows/uwp/enterprise/wip-hub) and use **WebView2 (Microsoft Edge Chromium-based)**. Add the WebView2 process, **msedgewebview2.exe**, to the protected apps list in your enterprise's WIP policy. An admin [adds this WIP policy through Intune](/mem/intune/apps/windows-information-protection-policy-create#to-add-a-wip-policy) with the following values.

- **Name**: Webview2
- **Publisher**: O=MICROSOFT CORPORATION, L=REDMOND, S=WASHINGTON, C=US
- **Product Name**: MICROSOFT EDGE WEBVIEW2
- **File**: MSEDGEWEBVIEW2.EXE
- **Min Version**: *
- **Max Version**: *

If the WIP policy hasn't been added, the add-in defaults to an older runtime. In the sections [Perpetual versions of Office on Windows](#perpetual-versions-of-office-on-windows) and [Microsoft 365 subscription versions of Office on Windows](#microsoft-365-subscription-versions-of-office-on-windows) earlier in this article, substitute **EdgeHTML (Microsoft Edge Legacy)** for **WebView2 (Microsoft Edge Chromium-based)** wherever the latter appears.

To determine if a document is WIP-protected, follow these steps.

1. Open the file.
1. Select the **File** tab on the ribbon.
1. Select **Info**.
1. In the upper section of the **Info** page, just below the file name, a WIP-enabled document will have a briefcase icon followed by **Managed by Work (...)**.

> [!NOTE]
> Support for WebView2 in WIP-enabled documents was added with build 16.0.16626.20132. If you're on an older build, your runtime defaults to **EdgeHTML (Microsoft Edge Legacy)**, regardless of policy.

## See also

- [Requirements for Running Office Add-ins](requirements-for-running-office-add-ins.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
