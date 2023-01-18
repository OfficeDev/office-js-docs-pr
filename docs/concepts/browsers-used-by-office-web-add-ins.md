---
title: Browsers used by Office Add-ins
description: Specifies how the operating system and Office version determine what browser is used by Office Add-ins.
ms.date: 01/18/2023
ms.localizationpriority: medium
---

# Browsers used by Office Add-ins

Office Add-ins are web applications that are displayed using iFrames when running in Office on the web. In Office for desktop and mobile clients, Office Add-ins use an embedded browser control (also known as a webview). Add-ins also need a JavaScript engine to run the JavaScript. Both the embedded browser and the engine are supplied by a browser installed on the user's computer.

Which browser is used depends on:

- The computer's operating system.
- Whether the add-in is running in Office on the web, in Office downloaded from a Microsoft 365 subscription, or in perpetual Office 2013 or later.
- Within the perpetual versions of Office on Windows, whether the add-in is running in the "retail" or "volume-licensed" variation.

> [!NOTE]
> This article assumes that the add-in is running in a document that is *not* protected with [Windows Information Protection (WIP)](/windows/uwp/enterprise/wip-hub). For WIP-protected documents, there are some exceptions to the information in this article. For more information, see [WIP-protected documents](#wip-protected-documents).

> [!IMPORTANT]
> **Internet Explorer and Microsoft Edge Legacy are still used in Office Add-ins**
>
> Some combinations of platforms and Office versions, including volume-licensed perpetual versions through Office 2019, still use the webview controls that come with Internet Explorer 11 and Microsoft Edge Legacy (EdgeHTML-based) to host add-ins, as explained in this article. We recommend (but don't require) that you continue to support these combinations, at least in a minimal way, by providing users of your add-in a graceful failure message when your add-in is launched in these browsers. Keep these additional points in mind:
>
> - Office on the web no longer opens in Internet Explorer or Microsoft Edge Legacy. Consequently, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) doesn't test add-ins in Office on the web on these browsers.
> - AppSource still tests for combinations of platform and Office *desktop* versions that use Internet Explorer and Microsoft Edge Legacy. However, it only issues a warning when the add-in doesn't support these browsers; the add-in isn't rejected by AppSource.
> - The [Script Lab tool](../overview/explore-with-script-lab.md) no longer supports Internet Explorer.
>
> For more information about supporting Internet Explorer and Microsoft Edge Legacy, including configuring a graceful failure message on your add-in, see [Support older Microsoft browsers and Office versions](../develop/support-ie-11.md).

The following sections specify which browser is used for the various platforms and operating systems.

## Non-Windows platforms

For these platforms, the platform alone determines the browser that is used.

|OS|Office version|Browser|
|:-----|:-----|:-----|
|any|Office on the web|The browser in which Office is opened.<br>(But note that Office on the web will not open in Internet Explorer.<br>Attempting to do so opens Office on the web in Edge.) |
|Mac|any|Safari with WKWebView|
|iOS|any|Safari with WKWebView|
|Android|any|Chrome|

> [!IMPORTANT]
> [Conditional Access](/azure/active-directory/conditional-access/overview) is not supported for Office Add-ins on iOS. Those add-ins use the Safari-based WKWebView, not an Edge-based browser control.

## Perpetual versions of Office on Windows

For perpetual versions of Office on Windows, the browser that is used is determined by the Office version, whether the license is retail or volume-licensed, and whether the Edge WebView2 (Chromium-based) is installed. The version of Windows doesn't matter, but note that Office Web Add-ins aren't supported on versions earlier than Windows 7 and Office 2021 isn't supported on versions earlier than Windows 10.

To determine whether Office 2016 or Office 2019 is retail or volume-licensed, use the format of the Office version and build number. (For Office 2013 and Office 2021, the distinction between volume-licensed and retail doesn't matter.)

- **Retail**: For both Office 2016 and 2019, the format is `YYMM (xxxxx.xxxxxx)`, ending with two blocks of five digits; for example, `2206 (Build 15330.20264`.
- **Volume-licensed**:
  - For Office 2016, the format is `16.0.xxxx.xxxxx`, ending with two blocks of *four* digits; for example, `16.0.5197.1000`.
  - For Office 2019, the format is `1808 (xxxxx.xxxxxx)`, ending with two blocks of *five* digits; for example, `1808 (Build 10388.20027)`. Note that the year and month is always `1808`.

| Office version | Retail vs. Volume-licensed | Edge WebView2 (Chromium-based) installed? | Browser |
|:-----|:-----|:-----|:-----|
| Office 2013 | Doesn't matter | Doesn't matter | Internet Explorer 11 |
| Office 2016 | Volume-licensed | Doesn't matter | Internet Explorer 11 |
| Office 2019 | Volume-licensed | Doesn't matter | Internet Explorer 11 |
| Office 2016 to Office 2019 | Retail | No | Microsoft Edge Legacy (EdgeHTML)<sup>1, 2</sup></br>If Edge isn't installed, Internet Explorer 11 is used. |
| Office 2016 to Office 2019 | Retail | Yes<sup>3</sup> | Microsoft Edge<sup>1</sup> with WebView2 (Chromium-based) |
| Office 2021 | Doesn't matter | Yes<sup>3</sup> | Microsoft Edge<sup>1</sup> with WebView2 (Chromium-based) |

<sup>1</sup> When you use Microsoft Edge, the Windows Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane. In Internet Explorer 11, the Narrator reads the title bar of the task pane, which comes from the **\<DisplayName\>** value in the add-in's manifest.

<sup>2</sup> If your add-in includes the **\<Runtimes\>** element in the manifest, then it will not use Microsoft Edge Legacy (EdgeHTML). If the conditions for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser. Otherwise, it uses Internet Explorer 11. For more information, see [Runtimes](/javascript/api/manifest/runtimes).

<sup>3</sup> On Windows versions prior to Windows 11, the WebView2 control must be installed so that Office can embed it. It's installed with perpetual Office 2021 or later; but it isn't automatically installed with Microsoft Edge. If you have an earlier version of perpetual Office, use the instructions for installing the control at [Microsoft Edge WebView2 / Embed web content ... with Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

## Microsoft 365 subscription versions of Office on Windows

For subscription Office on Windows, the browser that is used is determined by the operating system, the Office version, and whether the Edge WebView2 (Chromium-based) is installed.

|OS|Office version|Edge WebView2 (Chromium-based) installed?|Browser|
|:-----|:-----|:-----|:-----|
|Windows 7 | Microsoft 365| Doesn't matter | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver.&nbsp;<&nbsp;1903| Microsoft 365 | No| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp;16.0.11629<sup>2</sup>| Doesn't matter|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.11629&nbsp;*AND*&nbsp;<&nbsp;16.0.13530.20424<sup>2</sup>| Doesn't matter|Microsoft Edge Legacy (EdgeHTML)<sup>1, 3</sup>|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Window 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| No |Microsoft Edge Legacy (EdgeHTML)<sup>1, 3</sup>|
|Windows 8.1<br>Windows 10,<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| Yes<sup>4</sup>|  Microsoft Edge<sup>1</sup> with WebView2 (Chromium-based) |

<sup>1</sup> When you use Microsoft Edge, the Windows Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane. In Internet Explorer 11, the Narrator reads the title bar of the task pane, which comes from the **\<DisplayName\>** value in the add-in's manifest.

<sup>2</sup> See the [update history page](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) for more details.

<sup>3</sup> If your add-in includes the **\<Runtimes\>** element in the manifest, then it will not use Microsoft Edge Legacy (EdgeHTML). If the conditions for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser. Otherwise, it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version. For more information, see [Runtimes](/javascript/api/manifest/runtimes).

<sup>4</sup> On Windows versions prior to Windows 11, the WebView2 control must be installed so that Office can embed it. It's installed with Microsoft 365, Version 2101 or later, but it isn't automatically installed with Microsoft Edge. If you have an earlier version of Microsoft 365, use the instructions for installing the control at [Microsoft Edge WebView2 / Embed web content ... with Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/). On Microsoft 365 builds before 16.0.14326.xxxxx, you must also create the registry key **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** and set its value to `dword:00000001`.

## Working with Internet Explorer

Internet Explorer 11 doesn't support JavaScript versions later than ES5. If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you have two options.

- Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).
- Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.

For more information about these options, see [Support older Microsoft browsers and Office versions](../develop/support-ie-11.md).

Also, Internet Explorer 11 doesn't support some HTML5 features such as media, recording, and location. To learn more, see [Determine the browser the add-in is running in at runtime](../develop/support-ie-11.md#determine-the-browser-the-add-in-is-running-in-at-runtime).

## Troubleshoot Microsoft Edge issues

### Service Workers aren't working

Office Add-ins don't support Service Workers when the Microsoft Edge Legacy WebView, [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML), is used. They're supported with the [Chromium-based Edge WebView2](/microsoft-edge/hosting/webview2).

### Scroll bar doesn't appear in task pane

By default, scroll bars in Microsoft Edge are hidden until hovered over. To ensure that the scroll bar is always visible, the CSS styling that applies to the `<body>` element of the pages in the task pane should include the [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) property and it should be set to `scrollbar`.

### When debugging with the Microsoft Edge DevTools, the add-in crashes or reloads

Setting breakpoints in the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) can cause Office to think that the add-in is hung. It will automatically reload the add-in when this happens. To prevent this, add the following Registry key and value to the development computer: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### When the add-in tries to open, get "ADD-IN ERROR We can't open this add-in from the localhost" error

One known cause is that Microsoft Edge requires that localhost be given a loopback exemption on the development computer. Follow the instructions at [Cannot open add-in from localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### Get errors trying to download a PDF file

Directly downloading blobs as PDF files in an add-in isn't supported when Edge is the browser. The workaround is to create a simple web application that downloads blobs as PDF files. In your add-in, call the `Office.context.ui.openBrowserWindow(url)` method and pass the URL of the web application. This will open the web application in a browser window outside of Office.

## WIP-protected documents

Add-ins running in a [WIP-protected](/windows/uwp/enterprise/wip-hub) document never use **Microsoft Edge with WebView2 (Chromium-based)**. In the sections [Perpetual versions of Office on Windows](#perpetual-versions-of-office-on-windows) and [Microsoft 365 subscription versions of Office on Windows](#microsoft-365-subscription-versions-of-office-on-windows) earlier in this article, substitute **Microsoft Edge Legacy (EdgeHTML)** for **Microsoft Edge with WebView2 (Chromium-based)** wherever the latter appears.

To determine if a document is WIP-protected, follow these steps:

1. Open the file.
1. Select the **File** tab on the ribbon.
1. Select **Info**.
1. In the upper left of the **Info** page, just below the file name, a WIP-enable document will have briefcase icon followed by **Managed by Work (…)**.

## See also

- [Requirements for Running Office Add-ins](requirements-for-running-office-add-ins.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
