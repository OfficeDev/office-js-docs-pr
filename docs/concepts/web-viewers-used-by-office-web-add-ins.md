---
title: Web viewers used by Office Add-ins
description: ''
ms.date: 05/06/2019
localization_priority: Priority
---

# Web viewers used by Office Add-ins

Since Office Add-ins are web applications, they need a web page viewer to display the HTML pages of the web application and a JavaScript engine to run the JavaScript. Both are supplied by a browser installed on the user’s computer.

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
|Windows / non-subscription Office 2013 or later|Internet Explorer 11\*|
|Windows 10 ver. < 19*mm* / Office 365|Internet Explorer 11\*|
|Windows 10 ver. >= 19*mm* / Office 365 ver < 16.0.11629|Internet Explorer 11\*|
|Windows 10 ver. >= 19*mm* / Office 365 ver >= 16.0.11629|Edge|

\* Internet Explorer 11 does not support JavaScript versions later than ES5. If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill. Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.

> [!NOTE]
> Until they are generally available, you may need to be a Windows Insider to get a Windows version 19*mm* or greater, and you may need to be an Office Insider to get Office version 16.0.11629 or greater.
>
> To join Windows Insiders:
> 
> 1. Navigate to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.
> 2. You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows. Follow the instructions. When you select the pace of updates, choose the fastest option.
>
> To join Office Insiders
> 
> 1. Navigate to [Get started as an Office Insider](https://insider.office.com/en-us/join).
> 2. Follow the instruction on that page to join. When asked to specify a channel, select Insider.

## See also

- [Requirements for Running Office Add-ins](requirements-for-running-office-add-ins.md)
