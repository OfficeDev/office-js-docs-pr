---
title: Debug Office Add-ins on a Mac
description: Learn how to use a Mac to debug Office Add-ins.
ms.date: 02/04/2025
ms.localizationpriority: medium
---

# Debug Office Add-ins on a Mac

Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.

> [!IMPORTANT]
> Debugging add-ins with Office on Mac is only possible if Office is installed on the Mac from [Office.com](https://www.office.com), *not* the Apple App Store.

## Debugging with Safari Web Inspector on a Mac

If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.

To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version 16.9.1 (Build 18012504) or later. If you don't have an Office on Mac build, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).

To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > Mac App Store builds of Office don't support the `OfficeWebAddinDeveloperExtras` flag.

Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-mac.md). Right-click (or select and hold) the add-in and you should see an **Inspect Element** option in the context menu. Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.

> [!NOTE]
>
> - If you're debugging an [event-based](../develop/event-based-activation.md) or [spam-reporting](../outlook/spam-reporting.md) add-in in Outlook on Mac, follow the steps in [Debug event-based or spam-reporting add-ins](debug-autolaunch.md) after configuring the `OfficeWebAddinDeveloperExtras` property.
> - If you're trying to use the inspector and the dialog flickers, update Office to the latest version. If that doesn't resolve the flickering, try the following workaround.
>
>   1. Reduce the size of the dialog.
>   1. Choose **Inspect Element**, which opens in a new window.
>   1. Resize the dialog to its original size.
>   1. Use the inspector as required.

## Clearing the Office application's cache on a Mac

Add-ins are often cached in Office on Mac for performance reasons. For guidance on how to clear the Office cache on Mac, see [Clear the Office cache on Mac](clear-cache.md#clear-the-office-cache-on-mac).
