
# Set up an add-in catalog on Office 365

An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for SharePoint Add-ins and Office Add-ins. Administrators can upload Office Add-ins manifest files to the add-in catalog for use within their organization. When an administrator registers an add-in catalog as a trusted catalog (by setting group policy, or by specifying the trusted catalog on the  **Trusted Add-in Catalogs** tab of the **Options** dialog box by choosing **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**), users can insert the add-in from the insertion UI in an Office client application.

 >**Note**  The name "apps for Office" is changing to "Office Add-ins". During the transition, the documentation and the UI of some Office applications and Visual Studio tools might still use the term "app/apps". For details, see [New name for apps for Office and SharePoint: Office and SharePoint Add-ins](https://msdn.microsoft.com/en-us/library/fp161507.aspx#Anchor_2).


## To set up an add-in catalog in SharePoint Online


1. On the Office 365 admin center page, choose  **Admin**, and then choose  **SharePoint**.
    
2. In the left task pane, choose  **add-ins**.
    
3. On the  **add-ins** page, choose **Add-in Catalog**.
    
4. On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.
    
5. On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.
    
6. Specify the web site address.
    
7. Set the  **Storage Quota** to the lowest possible value (currently 110). You will only be installing add-in packages on this site collection and they are very small.
    
8. Set the  **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)
    
9. Choose  **OK**.
    
To add add-in to the Add-in Catalog Site, browse to the site you have just created. In the left navigation pane, choose  **Office Add-ins**, and then, to upload an Office Add-in manifest file, choose  **new add-in**.


## Additional resources



- [Publish your Office Add-in](../publish/publish.md)
    
- [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md)
    
