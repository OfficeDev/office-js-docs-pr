---
title: Debug add-ins using developer tools for Internet Explorer
description: Debug add-ins using the developer tools in Internet Explorer.
ms.date: 12/26/2024
ms.localizationpriority: medium
---

# Debug add-ins using developer tools in Internet Explorer

This article shows how to debug the client-side code (JavaScript or TypeScript) of your add-in when the following conditions are met.

- You cannot, or don't wish to, debug using tools built into your IDE; or you are encountering a problem that only occurs when the add-in is run outside the IDE.
- Your computer is using a combination of Windows and Office versions that use the Internet Explorer webview control, Trident.

To determine which browser or webview is being used on your computer, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> To install a version of Office that uses Trident or to force your current version to use Trident, see [Switch to the Trident webview](#switch-to-the-trident-webview).

## Debug a task pane add-in using the F12 tools

Windows 10 and 11 include a web development tool called "F12" because it was originally launched by pressing <kbd>F12</kbd> in Internet Explorer. F12 is now an independent application used to debug your add-in when it is running in the Internet Explorer webview control, Trident. The application is not available in earlier versions of Windows.

> [!NOTE]
> If your add-in has an [add-in command](../design/add-in-commands.md) that executes a function, the function runs in a hidden browser runtime process that the F12 tools cannot detect or attach to, so the technique described in this article cannot be used to debug code in the function.

The following steps are the instructions for debugging your add-in. If you just want to test the F12 tools themselves, see [Example add-in to test the F12 tools](#example-add-in-to-test-the-f12-tools).

1. [Sideload](test-debug-non-local-server.md) and run the add-in.
1. Launch the F12 development tools that corresponds to your version of Office.

   - For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe
   - For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe

   IEChooser opens with a window named **Choose target to debug**. Your add-in will appear in the window named by the filename of the add-in's home page. In the following screenshot, it is `Home.html`. Only processes that are running in Internet Explorer, or Trident, appear. The tool cannot attach to processes that are running in other browsers or webviews, including Microsoft Edge.

    :::image type="content" source="../images/choose-target-to-debug.png" alt-text="IEChooser screen, with several Internet Explorer and Trident processes listed. One is named Home.html.":::

1. Select your add-in's process; that is, its home page file name. This action will attach the F12 tools to the process and open the main F12 user interface.
1. Open the **Debugger** tab.
1. In the upper left of the tab, just below the debugger tool ribbon, there is a small folder icon. Select this to open a drop down list of the files in the add-in. The following is an example.

    :::image type="content" source="../images/f12-file-dropdown.png" alt-text="The upper left corner of debugger tab with a folder drop down open and a list of files.":::

1. Select the file that you want to debug and it opens in the the **script** (left) pane of the **Debugger** tab. If you're using a transpiler, bundler, or minifier, that changes the name of the file, it will have the final name that is actually loaded, not the original source file name.

1. Scroll to a line where you want to set a breakpoint and click in the margin to the left of the line number. You'll see a red dot to the left of the line and a corresponding line appears in the **Breakpoints** tab of the bottom right pane. The following screenshot is an example.

    :::image type="content" source="../images/debugger-home-js-02.png" alt-text="Debugger with breakpoint in home.js file.":::

1. Execute functions in the add-in as needed to trigger the breakpoint. When the breakpoint is hit, a right-pointing arrow appears on the red dot of the breakpoint. The following screenshot is an example.

    :::image type="content" source="../images/debugger-home-js-01.png" alt-text="Debugger with results from the triggered breakpoint.":::

> [!TIP]
> For more information about using the F12 tools, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).

### Example add-in to test the F12 tools

This example uses Word and a free add-in from AppSource.

1. Open Word and choose a blank document.
1. Select **Home** > **Add-ins**, then select **More Add-ins**.
1. In the **Office Add-ins** dialog, select the **STORE** tab.
1. Search for and select the **QR4Office** add-in. It opens in a task pane.
1. Launch the F12 development tools that corresponds to your version of Office as described in the preceding section.
1. In the F12 window, select **Home.html**.
1. In the **Debugger** tab, open the file **Home.js** as described in the preceding section.
1. Set the breakpoints on lines 310 and 312.
1. In the add-in, select the **Insert** button. One or the other breakpoint is hit.

## Debug a dialog in an add-in

If your add-in uses the Office Dialog API, the dialog runs in a separate process from the task pane (if any) and the tools must attach to that process. Follow these steps.

1. Run the add-in and the tools.
1. Open the dialog and then select the **Refresh** button in the tools. The dialog process is shown. Its name is the file name of the file that is open in the dialog.
1. Select the process to open it and debug just as described in the section [Debug a task pane add-in using the F12 tools](#debug-a-task-pane-add-in-using-the-f12-tools).

## Switch to the Trident webview

There are two ways to switch the Trident webview. You can run a simple command in a command prompt, or you can install a version of Office that uses Trident by default. We recommend the first method. But you should use the second in the following scenarios.

- Your project was developed with Visual Studio and IIS. It isn't Node.js based.
- You want to be absolutely robust in your testing.
- If for any reason the command line tool doesn't work.

### Switch via the command line

[!INCLUDE [Steps to switch webviews with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### Install a version of Office that uses Internet Explorer

[!INCLUDE [Steps to install Office that uses EdgeHTML (Edge Legacy) or Trident (Internet Explorer)](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## See also

- [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
