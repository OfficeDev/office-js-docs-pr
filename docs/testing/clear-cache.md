---
title: Clear the Office cache
description: Learn how to clear the Office cache on your computer.
ms.date: 01/29/2020
localization_priority: Priority
---

# Clear the Office cache

You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer. 

Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest. Doing so will allow Office to render the add-in as it's described by the updated manifest.

## Clear the Office cache on Windows

To remove all sideloaded add-ins from Excel, Word, and PowerPoint, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. 

To remove a sideloaded add-in from Outlook, use the steps outlined in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing) to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.

Additionally, to clear the Office cache on Windows 10 when the add-in is running in Microsoft Edge, you can use the Microsoft Edge DevTools.

> [!TIP]
> If you're just wanting the sideloaded add-in to reflect recent changes to its HTML or JavaScript source files, you shouldn't need to use the following steps to clear the cache. Instead, just put focus in the add-in's task pane (by clicking anywhere within the task pane) and then press **F5** to reload the add-in. 

> [!NOTE]
> To clear the Office cache using the following steps, your add-in must have a task pane. If your add-in is a UI-less add-in -- for example, one that uses the [on-send](/outlook/add-ins/outlook-on-send-addins) feature -- you'll need to add a task pane to your add-in that uses the same domain for [SourceLocation](../reference/manifest/sourcelocation.md), before you can use the following steps to clear the cache.

1. Install the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).

2. Open your add-in in Outlook.

3. Run the Microsoft Edge DevTools.

4. In the Microsoft Edge DevTools, open the **Local** tab. Your add-in will be listed by its name.

5. Select the add-in name to attach the debugger to your add-in. A new Microsoft Edge DevTools window will open when the debugger attaches to your add-in.

6. On the **Network** tab of the new window, select the **Clear cache** button.

    ![Microsoft Edge DevTools screenshot with the Clear cache button highlighted](../images/edge-devtools-clear-cache.png)

7. If completing these steps doesn't produce the desired result, you can also select the **Always refresh from server** button.

    ![Microsoft Edge DevTools screenshot with the Always refresh from server button highlighted](../images/edge-devtools-refresh-from-server.png)

## Clear the Office cache on Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  Clear the Office cache on iOS

To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.

## See also

- [Debug Office Add-ins](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [Debug your add-in with runtime logging](runtime-logging.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Office Add-ins XML manifest](../develop/add-in-manifests.md)
- [Validate an Office Add-in's manifest](troubleshoot-manifest.md)

