---
title: Sideload Office Add-ins on Mac for testing
description: Enable testing of your Office Add-in on Mac by sideloading the manifest to Word, Excel, or PowerPoint.
ms.date: 05/07/2026
ms.localizationpriority: medium
---

# Sideload Office Add-ins on Mac for testing

Sideload your add-in in Office on Mac to see how your add-in runs. For debugging, see [Debug Office Add-ins on a Mac](debug-office-add-ins-on-ipad-and-mac.md).

[!INCLUDE [Unified manifest note about platform sideloading restrictions](../includes/unified-manifest-sideload-restrictions-note.md)]

> [!NOTE]
> To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).

This article refers to sideloading Excel, PowerPoint, and Word add-ins that use the add-in only manifest. 

## Prerequisites

- A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products) installed.
- Word on Mac Version 15.18 (160109).
- Excel on Mac Version 15.19 (160206).
- PowerPoint on Mac Version 15.24 (160614).

## Sideload an add-in in Office on Mac

1. Open **Finder**.
1. Press <kbd>Cmd</kbd>+<kbd>Shift</kbd>+<kbd>G</kbd> to open the **Go to folder** dialog.
1. Enter one of the following filepaths, based on the application you want to use for sideloading. If the `wef` folder doesn't exist on your computer, create it.

    - For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

        > [!NOTE]
        > The remaining steps describe how to sideload a Word add-in, but they apply to Excel and PowerPoint too.

1. Copy your add-in's manifest file to this `wef` folder.

    :::image type="content" source="../images/all-my-files.png" alt-text="Wef folder in Office on Mac.":::

1. Open Word (or restart Word if it's already running), then open a document.
1. Select **Home** > **Add-ins**, and then select your add-in from the menu.
1. Verify that your add-in is displayed in Word.

    :::image type="content" source="../images/lorem-ipsum-wikipedia.png" alt-text="Office Add-in displayed in Office on Mac.":::

## Remove a sideloaded add-in

You can remove a previously sideloaded add-in by clearing the Office cache on your computer. See [Clear the Office cache](clear-cache.md).

## Next steps

- [Debug Office Add-ins on a Mac](debug-office-add-ins-on-ipad-and-mac.md)

## See also

- [Sideload Office Add-ins on iPad for testing](sideload-an-office-add-in-on-ipad.md)
- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Clear the Office cache](clear-cache.md)
- [Troubleshoot development errors](troubleshoot-development-errors.md)
