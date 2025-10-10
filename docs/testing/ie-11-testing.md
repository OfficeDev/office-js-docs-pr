---
title: Trident testing
description: Test your Office Add-in on the Trident webview associated with Internet Explorer 11.
ms.date: 09/30/2025
ms.localizationpriority: medium
---

# Test your Office Add-in on Trident

If you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control called "Trident" that's provided by Internet Explorer 11. You can use a command line to switch from a more modern webview used by add-ins to Trident for this testing. For information about which versions of Windows and Office use the Internet Explorer 11 webview control, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). In this article, "webview" refers to the combination of a webview control and a JavaScript engine.

> [!IMPORTANT]
> **Webviews from Internet Explorer and Microsoft Edge Legacy are still used in Office Add-ins**
>
> Some combinations of platforms and Office versions, including volume-licensed perpetual versions through Office 2019, still use the webview controls that come with Internet Explorer 11 (called "Trident") and Microsoft Edge Legacy (called "EdgeHTML") to host add-ins, as explained in [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Internet Explorer 11 was disabled in Windows 10 and Windows 11 in February 2023, and the UI for launching it was removed; but it's still installed on with those operating systems. So, Trident and other functionality from Internet Explorer can still be called programmatically by Office.
>
> We recommend (but don't require) that you support these combinations, at least in a minimal way, by providing users of your add-in a graceful failure message when your add-in is launched in these webviews. 

## Limitations, restrictions, and special considerations when working with Trident

When deciding whether, and how, to support Trident, keep these additional points in mind:

- Office on the web no longer opens in Internet Explorer or Microsoft Edge Legacy. Consequently, [Microsoft Marketplace](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center) doesn't test add-ins in Office on the web on these browsers.
- Microsoft Marketplace still tests for combinations of platform and Office *desktop* versions that use Trident or EdgeHTML. However, it only issues a warning when the add-in doesn't support these webviews; the add-in isn't rejected by Microsoft Marketplace.
- The [Script Lab tool](../overview/explore-with-script-lab.md) no longer supports Trident.
- Trident doesn't support JavaScript versions later than ES5. If you want to use the syntax and features of ECMAScript 2015 or later, you have to use a transpiler or polyfill or both. For more information about these options, see [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md).
- Trident doesn't support some HTML5 features such as media, recording, and location. To learn more, see [Determine the webview the add-in is running in at runtime](../develop/support-ie-11.md#determine-the-webview-the-add-in-is-running-in-at-runtime).
- Office on the web can't be opened in Internet Explorer 11, so you can't (and don't need to) test your add-in on Office on the web with Internet Explorer.
- Internet Explorer's Enhanced Security Configuration (ESC) must be turned off for Office Web Add-ins to work. If you're using a Windows Server computer as your client when developing add-ins, note that ESC is turned on by default in Windows Server.

## Switch to the Trident webview

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

There are two ways to switch the Trident webview. You can run a simple command in a command prompt, or you can install a version of Office that uses Trident by default. We recommend the first method, but you should use the second in the following scenarios.

- Your project was developed with Visual Studio and IIS. It isn't Node.js based.
- You want to be absolutely robust in your testing.
- You can't use the Beta channel for Microsoft 365 on your development computer.
- You're developing on a Mac.
- If for any reason the command line tool doesn't work.

### Switch via the command line

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### Install a version of Office that uses Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## See also

- [Test and debug Office Add-ins](test-debug-office-add-ins.md)
- [Sideload Office Add-ins for testing](test-debug-non-local-server.md)
- [Debug add-ins using developer tools for Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Attach a debugger from the task pane](attach-debugger-from-task-pane.md)
- [Runtimes in Office Add-ins](runtimes.md)
