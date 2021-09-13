---
title: Clear the Office cache
description: 'Learn how to clear the Office cache on your computer.'
ms.date: 08/02/2021
ms.localizationpriority: high
---

# Clear the Office cache

To remove an add-in that you've previously sideloaded on Windows, Mac, or iOS, you need to clear the Office cache on your computer.

Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using an updated manifest. Doing so allows Office to render the add-in as it's described by the updated manifest.

> [!NOTE]
> To remove a sideloaded add-in from Excel, OneNote, PowerPoint, or Word on the web, see [Sideload Office Add-ins in Office on the web for testing: Remove a sideloaded add-in](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in).

## Clear the Office cache on Windows

To remove all sideloaded add-ins from Excel, Word, and PowerPoint, delete the contents of the folder.

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

If the following folder exists, delete its contents too.

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

To remove a sideloaded add-in from Outlook, use the steps outlined in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in and then choose **Remove** to remove that specific add-in. If this add-in removal doesn't work, then delete the contents of the `Wef` folder as noted previously for Excel, Word, and PowerPoint.

Additionally, to clear the Office cache on Windows 10 when the add-in is running in Microsoft Edge, you can use the Microsoft Edge DevTools.

> [!TIP]
> If you only want the sideloaded add-in to reflect recent changes to its HTML or JavaScript source files, you shouldn't need to clear the cache. Instead, just put focus in the add-in's task pane (by clicking anywhere within the task pane) and then press **F5** to reload the add-in.

> [!NOTE]
> To clear the Office cache using the following steps, your add-in must have a task pane. If your add-in is a UI-less add-in -- for example, one that uses the [on-send](../outlook/outlook-on-send-addins.md) feature -- you'll need to add a task pane to your add-in that uses the same domain for [SourceLocation](../reference/manifest/sourcelocation.md), before you can use the following steps to clear the cache.

1. Install the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).

2. Open your add-in in the Office client.

3. Run the Microsoft Edge DevTools.

4. In the Microsoft Edge DevTools, open the **Local** tab. Your add-in will be listed by its name.

5. Select the add-in name to attach the debugger to your add-in. A new Microsoft Edge DevTools window will open when the debugger attaches to your add-in.

6. On the **Network** tab of the new window, select **Clear cache**.

    ![Microsoft Edge DevTools screenshot with the Clear cache button highlighted.](../images/edge-devtools-clear-cache.png)

7. If completing these steps doesn't produce the desired result, try selecting **Always refresh from server**.

    ![Microsoft Edge DevTools screenshot with the Always refresh from server button highlighted.](../images/edge-devtools-refresh-from-server.png)

## Clear the Office cache on Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## Clear the Office cache on iOS

To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, reinstall Office.

## See also

- [Troubleshoot development errors with Office Add-ins](troubleshoot-development-errors.md)
- [Debug Office Add-ins](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [Debug your add-in with runtime logging](runtime-logging.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Office Add-ins XML manifest](../develop/add-in-manifests.md)
- [Validate an Office Add-in's manifest](troubleshoot-manifest.md)
