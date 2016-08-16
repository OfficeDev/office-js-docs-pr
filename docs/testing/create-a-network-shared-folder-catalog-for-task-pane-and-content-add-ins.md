
# Sideload Office Add-ins for testing

You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share. 

>**Note:** To test an Office Add-in in Office Online, see [Sideload Office Add-ins in Office Online for testing](sideload-office-add-ins-for-testing.md). To test an add-in on an IPad or Mac, see [Sideload Office Add-ins on iPad and Mac for testing](sideload-an-office-add-in-on-ipad-and-mac.md ). To test an Outlook add-in, see [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md ).

Deploy only the manifest file to the shared folder catalog. Deploy the web application itself to a web server and specify the URL in the  **SourceLocation** element of the manifest file.

 >**Important:**  To help make add-ins that access external data and services more secure, your add-in should use a secure protocol such as Hypertext Transfer Protocol Secure (HTTPS) to connect to external data and services. You must use HTTPS if your add-in uses add-in commands.

## Share a folder

1. Navigate to the networked Windows computer where you want your add-in catalog, if it isn't the one you are at; and then navigate to the parent folder, or drive letter, of the folder you want to use as your add-in catalog.

2. Right-click the folder that you want as your catalog, and select **Properties**.

2. Open the **Sharing** tab.

3. On the **Choose people ...** page, add yourself and everyone else on your Office Add-in development team. If they are all members of a security group, you can just add the group. You will need at least **Read/Write** permission to the folder. 

4. Select **Share**.

5. Select **Done**.

6. Select **Close**.

Now you have shared the folder. In the next section you tell Office to trust the manifests inside it.


## Specify a file share as a trusted catalog

      
3. On your develoment computer, open a new document in Excel, Word, or PowerPoint.
    
4. Choose the  **File** tab, and then choose **Options**.
    
5. Choose  **Trust Center**, and then choose the  **Trust Center Settings** button.
    
6. Choose  **Trusted Add-in Catalogs**.
    
7. In the  **Catalog Url** box, enter the full network path to the catalog folder, and then choose **Add Catalog**.
    
8. Select the  **Show in Menu** check box, and then choose **OK**.

9. Close the Office application so your changes will take effect.
    
## Sideload add-ins


1. Put the manifest file of any add-in that you are testing into the catalog folder.

1. On your development computer, in Excel, Word, or PowerPoint, select  **My Add-ins** on the **Insert** tab of the ribbon.

2. Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.

4. Select the name of the add-in and select **OK** to insert the add-in.


## Additional resources

- [Use runtime logging to debug your manifest](../develop/use-runtime-logging-to-debug-manifest.md)
- [Publish your Office Add-in](../publish/publish.md)
    
