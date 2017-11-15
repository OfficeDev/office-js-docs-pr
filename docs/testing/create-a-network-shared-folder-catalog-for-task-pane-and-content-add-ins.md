---
title: Sideload Office Add-ins for testing
description: 
ms.date: 11/20/2017 
---

# Sideload Office Add-ins for testing

You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share. 

If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:

- [Sideload Office Add-ins in Office Online for testing](sideload-office-add-ins-for-testing.md)
- [Sideload Office Add-ins on iPad and Mac for testing](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md)

The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.

<iframe width="560" height="315" src="https://www.youtube.com/embed/XXsAw2UUiQo" frameborder="0" allowfullscreen></iframe>


## Share a folder

1. On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.

2. Open the context menu for the folder (right-click) and choose **Properties**.

3. Open the **Sharing** tab.

4. On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder. 

5. Choose **Share** > **Done** > **Close**.


## Specify the shared folder as a trusted catalog
      
1. Open a new document in Excel, Word, or PowerPoint.
    
2. Choose the **File** tab, and then choose **Options**.
    
3. Choose **Trust Center**, and then choose the  **Trust Center Settings** button.
    
4. Choose  **Trusted Add-in Catalogs**.
    
5. In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.
    
6. Select the **Show in Menu** check box, and then choose **OK**.

7. Close the Office application so your changes will take effect.
    

## Sideload your add-in

1. Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.

    > [!IMPORTANT]
    > To help make add-ins that access external data and services more secure, your add-in should use a secure protocol such as HTTPS to connect to external data and services. You must use HTTPS if your add-in uses add-in commands.

2. In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.

3. Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.

4. Select the name of the add-in and choose **OK** to insert the add-in.


## Additional resources

- [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md)
- [Publish your Office Add-in](../publish/publish.md)
    
