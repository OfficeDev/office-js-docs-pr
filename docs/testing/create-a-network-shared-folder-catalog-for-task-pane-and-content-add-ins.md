
# Sideload Office Add-ins for testing

You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share. 

>**Note:** To test an Office Add-in in Office Online, see [Sideload Office Add-ins in Office Online for testing](sideload-office-add-ins-for-testing.md). To test an add-in on an IPad or Mac, see [Sideload Office Add-ins on iPad and Mac for testing](sideload-an-office-add-in-on-ipad-and-mac.md ). To test an Outlook add-in, see [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md ).

Deploy only the manifest file to the shared folder catalog. Deploy the web application itself to a web server and specify the URL in the  **SourceLocation** element of the manifest file.

 >**Important:**  To help make add-ins that access external data and services more secure, your add-in should use a secure protocol such as Hypertext Transfer Protocol Secure (HTTPS) to connect to external data and services. You must use HTTPS if your add-in uses add-in commands.


## Specify a file share as a trusted catalog


1. Create a folder on a network share, for example  `\\MyShare\MyManifests`.
    
2. Put the manifest files for the task pane and content add-ins that you want to publish in this file share.
    
3. Open a new document in Excel, Word, or PowerPoint.
    
4. Choose the  **File** tab, and then choose **Options**.
    
5. Choose  **Trust Center**, and then choose the  **Trust Center Settings** button.
    
6. Choose  **Trusted Add-in Catalogs**.
    
7. In the  **Catalog Url** box, enter the path to the network share you created in Step 1, and then choose **Add Catalog**.
    
8. Select the  **Show in Menu** check box, and then choose **OK**.
    
After performing these steps, you can select  **My Add-ins** on the **Insert** tab of the ribbon and then choose **Shared Folder** at the top of the **Office Add-ins** dialog box to insert a task pane or content add-in from this catalog.


## Additional resources

- [Publish your Office Add-in](../publish/publish.md)
    
