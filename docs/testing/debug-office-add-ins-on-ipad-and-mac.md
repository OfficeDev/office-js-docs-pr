---
title: Debug Office Add-ins on a Mac
description: ''
ms.date: 11/26/2019
localization_priority: Priority
---

# Debug Office Add-ins on a Mac

Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.

## Debugging with Safari Web Inspector on a Mac

If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.

To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later. If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).

To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md). Right-click the add-in and you should see an **Inspect Element** option in the context menu. Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.

> [!NOTE]
> If you're trying to use the inspector and the dialog flickers, update Office to the latest version. If that doesn't resolve the flickering, try the following workaround:
> 1. Reduce the size of the dialog.
> 2. Choose **Inspect Element**, which opens in a new window.
> 3. Resize the dialog to its original size.
> 4. Use the inspector as required.

## Clearing the Office application's cache on a Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
