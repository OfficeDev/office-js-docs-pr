---
title: Debug Office Add-ins on a Mac
description: ''
ms.date: 04/24/2019
localization_priority: Priority
---

# Debug Office Add-ins on a Mac

You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.

## Debugging with Safari Web Inspector on a Mac

If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.

To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later. If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).

To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md). Right-click the add-in and you should see an **Inspect Element** option in the context menu.  Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.
> [!NOTE]
> If you are trying to use the inspector and the dialog flickers, try the following workaround:
> 1. Reduce the size of the dialog.
> 2. Choose **Inspect Element**, which opens in a new window.
> 3. Resize the dialog to its original size.
> 4. Use the inspector as required.

## Clearing the Office application's cache on a Mac or iPad

Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.

On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.

On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.
