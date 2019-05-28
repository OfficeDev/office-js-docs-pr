---
title: Browsers used by Office Add-ins
description: 'Specifies how the operating system and Office version determine what browser is used by Office Add-ins.'
ms.date: 05/28/2019
localization_priority: Priority
---

# Browsers used by Office Add-ins

Office add-ins are web applications that are displayed using iFrames when running in Office Online and using embedded browser controls in Office for desktop and mobile clients. Add-ins also need a JavaScript engine to run the JavaScript. Both the embedded browser and the engine are supplied by a browser installed on the user’s computer.

Which browser is used depends on:

- The computer’s operating system.
- Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.

The following table shows which browser is used for the various platforms and operating systems.

|**OS / Platform**|**Browser**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office Online|The browser in which Office Online is opened.|
|Mac|Safari|
|iOS|Safari|
|Android|Chrome|
|Windows / non-subscription Office 2013 or later|Internet Explorer 11|
|Windows 10 ver. < 1903 / Office 365|Internet Explorer 11|
|Windows 10 ver. >= 1903 / Office 365 ver < 16.0.11629|Internet Explorer 11|
|Windows 10 ver. >= 1903 / Office 365 ver >= 16.0.11629|Microsoft Edge\*|

\* When Microsoft Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane. When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.

> [!IMPORTANT]
> Internet Explorer 11 does not support JavaScript versions later than ES5. If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill. Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.

> [!NOTE]
> Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.
>
> To join Windows Insiders:
> 
> 1. Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.
> 2. You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows. Follow the instructions. When you select the pace of updates, choose the fastest option.
>
> To join Office Insiders:
> 
> 1. Go to [Get started as an Office Insider](https://insider.office.com/join).
> 2. Follow the instruction on that page to join. When asked to specify a channel, select Insider.

## See also

- [Requirements for Running Office Add-ins](requirements-for-running-office-add-ins.md)
