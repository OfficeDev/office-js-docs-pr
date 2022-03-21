---
title: Internet Explorer 11 testing
description: Test your Office Add-in on Internet Explorer 11.
ms.date: 11/02/2021
ms.localizationpriority: medium
---

# Test your Office Add-in on Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer still used in Office Add-ins**
>
> Microsoft is ending support for Internet Explorer, but this doesn't significantly affect Office Add-ins. Some combinations of platforms and Office versions, including one-time-purchase versions through Office 2019, will continue to use the webview control that comes with Internet Explorer 11 to host add-ins, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Moreover, support for these combinations, and hence for Internet Explorer, is still required for add-ins submitted to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Two things *are* changing:
>
> - Office on the web no longer opens in Internet Explorer. Consequently, AppSource no longer tests add-ins in Office on the web using Internet Explorer as the browser. But AppSource still tests for combinations of platform and Office *desktop* versions that use Internet Explorer.
> - The [Script Lab tool](../overview/explore-with-script-lab.md) no longer supports Internet Explorer.

If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11). You can use a command line to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing. For information about which versions of Windows and Office use the Internet Explorer 11 web view control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 does not support JavaScript versions later than ES5. If you want to use the syntax and features of ECMAScript 2015 or later, you have two options:
>
> - Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).
> - Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.
>
> For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).
>
> Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location. To learn more, see [Determine at runtime if the add-in is running in Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

> [!NOTE]
> Office on the web cannot be opened in Internet Explorer 11, so you cannot (and do not need to) test your add-in on Office on the web with Internet Explorer.

## Switch to the Internet Explorer 11 webview

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

There are two ways to switch the Internet Explorer webview. You can run a simple command in a command prompt, or you can install a version of Office that uses Internet Explorer by default. We recommend the first method. But you should use the second in the following scenarios.

- Your project was developed with Visual Studio and IIS. It isn't node.js-based.
- You want to be absolutely robust in your testing.
- If for any reason the command line tool doesn't work.

### Switch via the command line

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### Install a version of Office that uses Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## See also

* [Test and debug Office Add-ins](test-debug-office-add-ins.md)
* [Sideload Office Add-ins for testing](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Debug add-ins using developer tools for Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
* [Attach a debugger from the task pane](attach-debugger-from-task-pane.md)
