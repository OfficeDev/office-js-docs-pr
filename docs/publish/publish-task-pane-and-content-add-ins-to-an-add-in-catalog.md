---
title: Publish task pane and content add-ins to a SharePoint catalog
description: To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization.
ms.date: 01/23/2018
---

# Publish task pane and content add-ins to a SharePoint catalog

An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.

> [!IMPORTANT]
> - Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.
> - If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.
> - SharePoint catalogs are not supported for Office for Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).   

## Set up an add-in catalog

Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.

### To set up an add-in catalog for on-premises SharePoint

> [!NOTE]
> The UI in on-premises SharePoint still refers to add-ins as **apps**.

1. Browse to the  **Central Administration Site**.
    
2. In the left task pane, choose  **Apps**.
    
3. On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.
    
4. On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.
    
5. Choose  **View site settings**.
    
6. On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.
    
7. To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.
    
8. In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.
    
9. To add an add-in to the Office Add-ins add-in catalog, choose **Apps for Office**.

### To set up an add-in catalog on Office 365

1. On the Office 365 admin center page, choose  **Admin**, and then choose  **SharePoint**.
    
2. In the left task pane, choose  **add-ins**.
    
3. On the  **add-ins** page, choose **Add-in Catalog**.
    
4. On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.
    
5. On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.
    
6. Specify the web site address.
    
7. Set the  **Storage Quota** to the lowest possible value (currently 110). You will only be installing add-in packages on this site collection and they are very small.
    
8. Set the  **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)
    
9. Choose  **OK**.
    
10. To add an add-in to the Add-in Catalog Site, browse to the site you have just created. In the left navigation pane, choose  **Office Add-ins**, and then, to upload an Office Add-in manifest file, choose  **new add-in**.

## Publish an add-in to an add-in catalog

To publish an add-in to an add-in catalog, complete the following steps.

1. Browse to the add-in catalog:

	- Open the SharePoint Central Administration main page.
	
	- Select  **Add-ins**.
	
	- Select  **Manage Add-in Catalog**.
	
	- Choose the link provided, and then choose  **Office Add-ins** on the left navigation bar.
    
2. Choose the  **Click to add new item** link.
    
3. Choose  **Browse**, and then specify the [manifest](../develop/add-in-manifests.md) to upload.
    
    Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.

## End user experience with the add-in catalog

End users can access the add-in catalog in an Office application by completing the following steps:

1. In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.
    
2. Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog. 
    
    For example, if the URL of the Office Add-ins catalog is:
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    Specify just the URL of the parent site collection:
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.

Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy. For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).
