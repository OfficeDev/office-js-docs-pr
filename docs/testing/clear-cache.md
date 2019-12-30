---
title: Clear the Office cache
description: Learn how to clear the Office cache on your computer.
ms.date: 12/31/2019
localization_priority: Priority
---

# Clear the Office cache

You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer. 

Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest. Doing so will allow Office to render the add-in as it's described by the updated manifest.

## Clear the Office cache on Windows

To clear the Office cache on Windows, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

## Clear the Office cache on Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  Clear the Office cache on iOS

To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.

## See also

- [Office Add-ins XML manifest](../develop/add-in-manifests.md)
- [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md)
- [Debug your add-in with runtime logging](runtime-logging.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Debug Office Add-ins](debug-add-ins-using-f12-developer-tools-on-windows-10.md)