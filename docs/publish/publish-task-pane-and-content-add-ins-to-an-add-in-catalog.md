
# Publish task pane and content add-ins to an add-in catalog on SharePoint
After you upload your add-in manifest to a catalog, users can configure their Office clients so your add-in is available from the Office Add-ins dialog box.


Use the following steps to upload the manifest for your task pane or content add-in to an Office Add-ins catalog on SharePoint. 

 >**Note**  The name "apps for Office" is changing to "Office Add-ins". During the transition, the documentation and the UI of some Office applications and Visual Studio tools might still use the term "app/apps". For details, see [New name for apps for Office and SharePoint: Office and SharePoint Add-ins](https://msdn.microsoft.com/en-us/library/fp161507.aspx#Anchor_2).


## Create or find the add-in catalog

To set up an add-in catalog, see [Set up an add-in catalog on SharePoint](../publish/set-up-an-add-in-catalog-on-sharepoint.md) or [Set up an add-in catalog on Office 365](../publish/set-up-an-add-in-catalog-on-office-365.md).

If an add-in catalog is already set up for a SharePoint web application, to find it:


1. Open the SharePoint Central Administration main page.
    
2. Select  **Add-ins**.
    
3. Select  **Manage Add-in Catalog**.
    
4. Choose the link provided, and then choose  **Office Add-ins** on the left navigation bar.
    

## Publish to an add-in catalog


1. Browse to the add-in catalog.
    
2. Choose the  **Click to add new item** link.
    
3. Choose  **Browse**, and then specify the [manifest](../../docs/overview/add-in-manifests.md) to upload.
    
    Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.
    
After you upload add-in manifests to the Office Add-ins catalog, users can access the add-ins by doing the following:


1. In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.
    
2. Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog. For example, if the URL of the Office Add-ins catalog is:
    
    https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog
    
    Specify just the URL of the parent site collection:
    
    https:// _domain_ /sites/ _AddinCatalogSiteCollection_
    
3. Close and reopen the Office application. The add-in catalog will be available in the  **Office Add-ins** dialog box.
    
Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy. For details, see the section "Using Group Policy to manage how users can install and use Office Add-ins" in [Overview of Office Add-ins](https://technet.microsoft.com/en-us/library/jj219429.aspx) on TechNet.


## Additional resources


- [Publish your Office Add-in](../publish/publish.md)
    
- [Set up an add-in catalog on SharePoint](../publish/set-up-an-add-in-catalog-on-sharepoint.md)
    
- [Set up an add-in catalog on Office 365](../publish/set-up-an-add-in-catalog-on-office-365.md)
    
