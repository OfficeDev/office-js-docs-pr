
# Set up an add-in catalog on SharePoint

An add-in catalog is a document library on SharePoint where manifest files for task pane and content Office Add-ins, as well as SharePoint Add-ins, can be published. For Office Add-ins, an administrator uploads a [manifest file](../../docs/overview/add-in-manifests.md) to the add-in catalog. When an administrator registers an add-in catalog as a trusted catalog (by setting group policy, or from the **Trusted Add-ins Catalog** page of the **Options** dialog box, choosing **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**), users can insert the add-in from the insertion UI in an Office client application.

 >**Note**  The name "apps for Office" is changing to "Office Add-ins". During the transition, the documentation and the UI of some Office applications and Visual Studio tools might still use the term "app/apps". For details, see [New name for apps for Office and SharePoint: Office and SharePoint Add-ins](https://msdn.microsoft.com/en-us/library/fp161507.aspx#Anchor_2).

Only one add-in catalog for Office Add-ins can exist per SharePoint web application. To set up the add-in catalog for a web application:

1. Browse to the  **Central Administration Site** ( **Start** > **All Programs** > **Microsoft SharePoint 2013 Products** > **SharePoint 2013 Central Administration**).
    
2. In the left task pane, choose  **Add-ins**.
    
3. On the  **Add-ins** page, under **Add-in Management**, choose  **Manage Add-in Catalog**.
    
4. On the  **Manage Add-in Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.
    
5. Choose  **View site settings**.
    
6. On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.
    
7. To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.
    
8. In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.
    
9. To add add-ins to the Office Add-ins add-in catalog, choose  **Office Add-ins**.
    

## Additional resources


- [Set up an add-in catalog on SharePoint Online](../../docs/publish/set-up-an-add-in-catalog-on-office-365.md)
    
