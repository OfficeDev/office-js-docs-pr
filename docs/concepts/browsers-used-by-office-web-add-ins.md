---
title: Browsers used by Office Add-ins
description: 'Specifies how the operating system and Office version determine what browser is used by Office Add-ins.'
ms.date: 06/01/2020
localization_priority: Normal
---

# Browsers used by Office Add-ins

Office add-ins are web applications that are displayed using iFrames when running in Office on the web and using embedded browser controls in Office for desktop and mobile clients. Add-ins also need a JavaScript engine to run the JavaScript. Both the embedded browser and the engine are supplied by a browser installed on the user's computer.

Which browser is used depends on:

- The computer's operating system.
- Whether the add-in is running in Office on the web, Office 365, or non-subscription Office 2013 or later.

The following table shows which browser is used for the various platforms and operating systems.

|**OS / Platform**|**Office version**|**Browser**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|any|Office on the web|The browser in which Office is opened.|
|Mac|any|Safari|
|iOS|any|Safari|
|Android|any|Chrome|
|Windows 7, 8.1, 10 | non-subscription Office 2013 or later|Internet Explorer 11|
|Windows 8.1, 10 ver.&nbsp;<&nbsp;1903 | Office 365|Internet Explorer 11<br>unless Edge WebView2 is installed (see below)|
|Windows 10 ver.&nbsp;>=&nbsp;1903 | Office 365 ver.&nbsp;<&nbsp;16.0.11629<sup>1</sup>|Internet Explorer 11<br>unless Edge WebView2 is installed (see below)|
|Windows 10 ver.&nbsp;>=&nbsp;1903 | Office 365 ver.&nbsp;>=&nbsp;16.0.11629&nbsp;_AND_&nbsp;<&nbsp;16.0.12822<sup>1</sup>|Microsoft Edge<sup>2, 3</sup> with original WebView, also called EdgeHTML<br>unless Edge WebView2 is installed (see below)|
|Windows 8.1, 10 | Office 365 ver.&nbsp;>=&nbsp;16.0.12822<sup>1</sup> |Microsoft Edge<sup>2, 3</sup> with WebView2 &mdash; Chromium based Edge web control &mdash; if it is installed.<br>If it is not, Microsoft Edge with original Webview, also called EdgeHTML.

<sup>1</sup> See the [update history page](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) for more details.

<sup>2</sup> When Microsoft Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane. When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.

<sup>3</sup> If your add-in includes the `Runtimes` element in the manifest, it uses Internet Explorer 11 regardless of the Windows or Office 365 version. For more information, see [Runtimes](../reference/manifest/runtimes.md).

> [!IMPORTANT]
> Internet Explorer 11 does not support JavaScript versions later than ES5. If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill. Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.

## Troubleshooting Microsoft Edge Issues

### Service Workers are not working

Office Add-ins do not support Service Workers on [Microsoft Edge WebView](/microsoft-edge/hosting/webview). Please see the [Office Add-ins overview](../overview/office-add-ins.md) for the latest supported features on the Edge WebView control. We are working hard to bring the new [Chromium-based Edge WebView2](/microsoft-edge/hosting/webview2) to the Office Add-ins platform, which we expect will support Service Workers.

### Chromium-based Edge is installed on my development computer, but my add-in does not use it

The base browser in [Microsoft Edge](https://support.microsoft.com/help/4501095/download-the-new-microsoft-edge-based-on-chromium) has changed to Chromium. The older base, called EdgeHTML, is not removed when Chromium-based Edge is installed. Office will still use the EdgeHTML base for add-ins until a build of Office 365 that supports Chromium is installed on the computer. We expect these builds to ship in 2020. They will likely appear in the Insiders channel in the first half of the year.

### Scroll bar does not appear in task pane

By default, scroll bars in Microsoft Edge are hidden until hovered over. To ensure that the scroll bar is always visible, the CSS styling that applies to the `<body>` element of the pages in the task pane should include the [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) property and it should be set to `scrollbar`. 

### When debugging with the Microsoft Edge DevTools, the add-in crashes or reloads

Setting breakpoints in the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) can cause Office to think that the add-in is hung. It will automatically reload the add-in when this happens. To prevent this, add the following Registry key and value to the development computer: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### When the add-in tries to open, get "ADD-IN ERROR We can't open this add-in from the localhost" error

One known cause is that Microsoft Edge requires that localhost be given a loopback exemption on the development computer. Follow the instructions at [Cannot open add-in from localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### Get errors trying to download a PDF file

Directly downloading blobs as PDF files in an add-in is not supported when Edge is the browser. The workaround is to create a simple web application that downloads blobs as PDF files. In your add-in call the `Office.context.ui.openBrowserWindow(url)` method and pass the URL of the web application. This will open the web application in a browser window outside of Office.

## See also

- [Requirements for Running Office Add-ins](requirements-for-running-office-add-ins.md)
