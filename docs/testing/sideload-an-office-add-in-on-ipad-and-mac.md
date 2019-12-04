---
title: Sideload Office Add-ins on iPad and Mac for testing
description: ''
ms.date: 11/26/2019
localization_priority: Priority
---

# Sideload Office Add-ins on iPad and Mac for testing

To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately. 

## Prerequisites for Office on iOS

- A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.
    
- An iPad running iOS 8.2 or later with [Excel on iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.
    
- The manifest .xml file for the add-in you want to test.
    

## Prerequisites for Office on Mac

- A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.
    
- Word on Mac version 15.18 (160109).
   
- Excel on Mac version 15.19 (160206).

- PowerPoint on Mac version 15.24 (160614)
    
- The manifest .xml file for the add-in you want to test.
    

## Sideload an add-in on Excel or Word on iPad

1. Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with  **Trust This Computer?**. Choose **Trust** to continue.

2. In iTunes, choose the  **iPad** icon below the menu bar.

3. Under  **Settings** on the left side of iTunes, choose **Apps**.

4. On the right side of iTunes, scroll down to  **File Sharing**, and then choose  **Excel** or **Word** in the **Add-ins** column.

5. At the bottom of the  **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload. 
    
6. Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the  **Home** button, and then close and restart the app.
    
7. Open a document.
    
8. Choose  **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.
    
    ![Insert Add-ins in the Excel app](../images/excel-insert-add-in.png)


## Sideload an add-in in Office on Mac

> [!NOTE]
> To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).

1. Open  **Terminal** and go to one of the following folders where you'll save your add-in's manifest file. If the `wef` folder doesn't exist on your computer, create it.
    
    - For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`    
    - For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`
    
2. Open the folder in  **Finder** using the command `open .` (including the period or dot). Copy your add-in's manifest file to this folder.
    
    ![Wef folder in Office on Mac](../images/all-my-files.png)

3. Open Word, and then open a document. Restart Word if it's already running.
    
4. In Word, choose  **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.
    
    ![My Add-ins in Office on Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu. 
    
5. Verify that your add-in is displayed in Word.
    
    ![Office Add-in displayed in Office on Mac](../images/lorem-ipsum-wikipedia.png)
    
### Clearing the Office application's cache on a Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## See also

- [Debug Office Add-ins on iPad and Mac](debug-office-add-ins-on-ipad-and-mac.md)
