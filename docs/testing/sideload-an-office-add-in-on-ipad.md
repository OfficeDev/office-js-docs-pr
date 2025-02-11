---
title: Sideload Office Add-ins on iPad for testing
description: Test your Office Add-in on iPad by sideloading.
ms.date: 02/12/2025
ms.localizationpriority: medium
---

# Sideload Office Add-ins on iPad for testing

To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.

[!INCLUDE [Unified manifest note about platform sideloading restrictions](../includes/unified-manifest-sideload-restrictions-note.md)]

> [!NOTE]
> To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).

## Prerequisites for Office on iOS

- A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.
  > [!IMPORTANT]
  > If you're running macOS Catalina, [iTunes is no longer available](https://support.apple.com/119585), so you should follow the instructions in the section [Sideload an add-in on Excel or Word on iPad using macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) later in this article.

- An iPad running iOS 8.2 or later with [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) or [Word](https://apps.apple.com/app/microsoft-word/id586447913) installed, and a sync cable.

- The **manifest.xml** file for the add-in you want to test.

> [!IMPORTANT]
> You can't sideload an add-in on an iPad when it's running on localhost. You must *first* [deploy the add-in](../publish/publish.md) to a non-local web server or web service. Then, update all the URLs in the manifest to point to the new domain. 

## Sideload an add-in on Excel or Word on iPad using iTunes

1. Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**. Choose **Trust** to continue.

1. In iTunes, choose the **iPad** icon below the menu bar.

1. Under **Settings** on the left side of iTunes, choose **Apps**.

1. On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.

1. At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest.xml file of the add-in you want to sideload.

1. Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.

1. Open a document.

1. Select **Home** > **Add-ins**. In the dialog that opens, select **See all** under the **My Add-ins** heading, then select your sideloaded add-in from the array under **My Add-ins**.

## Sideload an add-in on Excel or Word on iPad using macOS Catalina

> [!IMPORTANT]
> With the introduction of macOS Catalina, Apple discontinued iTunes on Mac and integrated functionality required to sideload apps into **Finder**. To learn more, see [Use the Finder to share files between your Mac and your iPhone, iPad, iPod touch](https://support.apple.com/119585).

1. Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**. Choose **Trust** to continue. You may also be asked if this is a new iPad or if you're restoring one.

1. In Finder, under **Locations**, choose the **iPad** icon below the menu bar.

1. On the top of the Finder window, click on **Files**, and then locate **Excel** or **Word**.

1. From a different Finder window, drag and drop the manifest.xml file of the add-in you want to side load onto the **Excel** or **Word** file in the first Finder window.

1. Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.

1. Open a document.

1. Select **Home** > **Add-ins**. In the dialog that opens, select **See all** under the **My Add-ins** heading, then select your sideloaded add-in from the array under **My Add-ins**.

## Remove a sideloaded add-in

You can remove a previously sideloaded add-in by clearing the Office cache on your computer. Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).

## See also

- [Sideload Office Add-ins on Mac for testing](sideload-an-office-add-in-on-mac.md)
- [Debug Office Add-ins on a Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)
