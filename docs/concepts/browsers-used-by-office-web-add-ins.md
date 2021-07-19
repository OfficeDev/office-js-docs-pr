---
title: Browsers used by Office Add-ins
description: 'Specifies how the operating system and Office version determine what browser is used by Office Add-ins.'
ms.date: 06/18/2021
localization_priority: Normal
---

# Browsers used by Office Add-ins

Office Add-ins are web applications that are displayed using iFrames when running in Office on the web and using embedded browser controls in Office for desktop and mobile clients. Add-ins also need a JavaScript engine to run the JavaScript. Both the embedded browser and the engine are supplied by a browser installed on the user's computer.

Which browser is used depends on:

- The computer's operating system.
- Whether the add-in is running in Office on the web, Microsoft 365, or non-subscription Office 2013 or later.

> [!IMPORTANT]
> **Internet Explorer still used in Office Add-ins**
>
> Microsoft is ending support for Internet Explorer, but this doesn't significantly affect Office Add-ins. Some combinations of platforms and Office versions, including all one-time-purchase versions through Office 2019, will continue to use the webview control that comes with Internet Explorer 11 to host add-ins, as explained in this article. Moreover, support for these combinations, and hence for Internet Explorer, is still required for add-ins submitted to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Two things *are* changing:
>
> - AppSource no longer tests add-ins in Office on the web using Internet Explorer as the browser. But AppSource still tests for combinations of platform and Office *desktop* versions that use Internet Explorer.
> - The [Script Lab tool](../overview/explore-with-script-lab.md) will stop working in Internet Explorer sometime in 2021.

The following table shows which browser is used for the various platforms and operating systems.

|OS|Office version|Edge WebView2 (Chromium-based) installed?|Browser|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|any|Office on the web|Not applicable|The browser in which Office is opened.|
|Mac|any|Not applicable|Safari|
|iOS|any|Not applicable|Safari|
|Android|any|Not applicable|Chrome|
|Windows 7, 8.1, 10 | non-subscription Office 2013 or later|Doesn't matter|Internet Explorer 11|
|Windows 7 | Microsoft 365| Doesn't matter | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver.&nbsp;<&nbsp;1903| Microsoft 365 | No| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903 | Microsoft 365 ver.&nbsp;<&nbsp;16.0.11629<sup>1</sup>| Doesn't matter|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.11629&nbsp;_AND_&nbsp;<&nbsp;16.0.13530.20424<sup>1</sup>| Doesn't matter|Microsoft Edge<sup>2, 3</sup> with original WebView (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp;1903 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>1</sup>| No |Microsoft Edge<sup>2, 3</sup> with original WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>1</sup>| Yes<sup>4</sup>|  Microsoft Edge<sup>2</sup> with WebView2 (Chromium-based) |

<sup>1</sup> See the [update history page](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) for more details.

<sup>2</sup> When Microsoft Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane. When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.

<sup>3</sup> If your add-in includes the `<Runtimes>` element in the manifest, then it will not use Microsoft Edge with the original WebView (EdgeHTML). If the conditions for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser. Otherwise, it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version. For more information, see [Runtimes](../reference/manifest/runtimes.md).

<sup>4</sup> The embeddable WebView2 control must be installed in addition to the installation of Microsoft Edge so that Office can embed it. To install it, see [Microsoft Edge WebView2 / Embed web content ... with Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

> [!IMPORTANT]
> Internet Explorer 11 does not support JavaScript versions later than ES5. If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you have two options:
>
> - Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).
> - Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.
>
> For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).
>
> Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.

## Troubleshooting Microsoft Edge issues

### Service Workers are not working

Office Add-ins do not support Service Workers when the original Microsoft Edge WebView, [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML), is used. They are supported with the [Chromium-based Edge WebView2](/microsoft-edge/hosting/webview2).

### Scroll bar does not appear in task pane

By default, scroll bars in Microsoft Edge are hidden until hovered over. To ensure that the scroll bar is always visible, the CSS styling that applies to the `<body>` element of the pages in the task pane should include the [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) property and it should be set to `scrollbar`.

### When debugging with the Microsoft Edge DevTools, the add-in crashes or reloads

Setting breakpoints in the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) can cause Office to think that the add-in is hung. It will automatically reload the add-in when this happens. To prevent this, add the following Registry key and value to the development computer: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### When the add-in tries to open, get "ADD-IN ERROR We can't open this add-in from the localhost" error

One known cause is that Microsoft Edge requires that localhost be given a loopback exemption on the development computer. Follow the instructions at [Cannot open add-in from localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### Get errors trying to download a PDF file

Directly downloading blobs as PDF files in an add-in is not supported when Edge is the browser. The workaround is to create a simple web application that downloads blobs as PDF files. In your add-in, call the `Office.context.ui.openBrowserWindow(url)` method and pass the URL of the web application. This will open the web application in a browser window outside of Office.

## See also

- [Requirements for Running Office Add-ins](requirements-for-running-office-add-ins.md)
