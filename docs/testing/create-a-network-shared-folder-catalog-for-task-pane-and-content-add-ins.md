---
title: Sideload Office Add-ins for testing
description: ''
ms.date: 03/19/2019
localization_priority: Priority
---

# Sideload Office Add-ins for testing

You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).

> [!NOTE]
> If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you. For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).

This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows. If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:

- [Sideload Office Add-ins in Office for testing](sideload-office-add-ins-for-testing.md)
- [Sideload Office Add-ins on iPad and Mac for testing](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing)


The following video walks you through the process of sideloading your add-in on Office using a shared folder catalog.  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## Share a folder

1. In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.

2. Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.

3. Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.

    ![folder Properties dialog with the Sharing tab and Share button highlighted](../images/sideload-windows-properties-dialog.png)

4. Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in. You will need at least **Read/Write** permission to the folder. After you have finished choosing people to share with, choose the **Share** button.

5. When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name. (You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.

   ![Network access dialog with the share path highlighted](../images/sideload-windows-network-access-dialog.png)

6. Choose the **Close** button to close the **Properties** dialog window.

## Specify the shared folder as a trusted catalog

1. Open a new document in Excel, Word, or PowerPoint.

2. Choose the **File** tab, and then choose **Options**.

3. Choose **Trust Center**, and then choose the **Trust Center Settings** button.

4. Choose **Trusted Add-in Catalogs**.

5. In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously. If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.

    ![folder Properties dialog with the Sharing tab and network path highlighted](../images/sideload-windows-properties-dialog-2.png)

6. After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.

7. Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.

    ![Trust Center dialog with catalog selected](../images/sideload-windows-trust-center-dialog.png)

8. Choose the **OK** button to close the **Word Options** dialog window.

9. Close and reopen the Office application so your changes will take effect.


## Sideload your add-in


1. Put the manifest XML file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.

3. Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.

4. Select the name of the add-in and choose **OK** to insert the add-in.


## See also

- [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md)
- [Publish your Office Add-in](../publish/publish.md)
