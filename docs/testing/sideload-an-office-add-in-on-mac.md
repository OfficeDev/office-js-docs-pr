---
title: Sideload Office Add-ins on Mac for testing
description: Test your Office Add-in on Mac by sideloading.
ms.date: 02/12/2025
ms.localizationpriority: medium
---

# Sideload Office Add-ins on Mac for testing

To see how your add-in will run on Office on Mac, you can sideload your add-in's manifest. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.

[!INCLUDE [Unified manifest note about platform sideloading restrictions](../includes/unified-manifest-sideload-restrictions-note.md)]

> [!NOTE]
> To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).

## Prerequisites for Office on Mac

- A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products) installed.

- Word on Mac Version 15.18 (160109).

- Excel on Mac Version 15.19 (160206).

- PowerPoint on Mac Version 15.24 (160614).

- The manifest .xml file for the add-in you want to test.

## Sideload an add-in in Office on Mac

1. Use **Finder** to sideload the manifest file. Open **Finder** and then enter <kbd>Cmd</kbd>+<kbd>Shift</kbd>+<kbd>G</kbd> to open the **Go to folder** dialog.

1. Enter one of the following filepaths, based on the application you want to use for sideloading. If the `wef` folder doesn't exist on your computer, create it.

    - For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

        > [!NOTE]
        > The remaining steps describe how to sideload a Word add-in.

1. Copy your add-in's manifest file to this `wef` folder.

    ![Wef folder in Office on Mac.](../images/all-my-files.png)

1. Open Word (or restart Word if it's already running), then open a document.

1. Select **Home** > **Add-ins**, and then select your add-in from the menu.

1. Verify that your add-in is displayed in Word.

    ![Office Add-in displayed in Office on Mac.](../images/lorem-ipsum-wikipedia.png)

## Remove a sideloaded add-in

You can remove a previously sideloaded add-in by clearing the Office cache on your computer. Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).

## See also

- [Sideload Office Add-ins on iPad for testing](sideload-an-office-add-in-on-ipad.md)
- [Debug Office Add-ins on a Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)
