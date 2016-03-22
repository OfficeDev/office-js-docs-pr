
# Create a network shared folder catalog for task pane and content add-ins


A shared folder catalog provides a way to publish the manifests for task pane and content Office Add-ins to a network file share. Users can then get add-ins by specifying this file share as a trusted catalog, using the steps in the following procedure.

The manifest file is an XML file that enables you to declaratively describe how your add-in should be activated when an end user installs and uses it with Office documents and applications. For more information, see [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md).

The manifest file is the only file that you should deploy to the shared folder catalog. Deploy the web application itself to a web server and specify the URL in the  **SourceLocation** element of the manifest file.

 >**Important**  To help make add-ins that access external data and services more secure, your add-in should use a secure protocol such as Hypertext Transfer Protocol Secure (HTTPS) to connect to external data and services.


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

Any additional manifest files you put in this file share will be available to users that have specified this shared folder catalog.


## Additional resources



- [Publish your Office Add-in](../publish/publish.md)
    
