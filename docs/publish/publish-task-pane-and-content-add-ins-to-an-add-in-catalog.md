---
title: Publish task pane and content add-ins to a SharePoint catalog
description: To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization.
ms.date: 03/19/2019
localization_priority: Priority
---

# Publish task pane and content add-ins to a SharePoint catalog

An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.

> [!IMPORTANT]
> - Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.
> - If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.
> - SharePoint catalogs are not supported for Office for Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).   

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

### To set up an app catalog on Office 365

Even though SharePoint names the catalog an "app" catalog, you can register Office Add-ins in the app catalog.

1. On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.

    > [!NOTE]
    > You need to use the Classic SharePoint admin center to set up the catalog. If you are in the new SharePoint admin center, choose **Class SharePoint admin center** in the left pane.

2. In the left task pane, choose  **apps**.

3. On the **apps** page, choose **App Catalog**.

4. On the **App Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.

5. On the **Create App Catalog Site Collection** page, specify the title of your App Catalog site.

6. Specify the **Web Site Address**.

7. Specify an **Administrator**.

8. Set the **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your app catalog site.)

9. Choose **OK**.

To add an add-in to the App Catalog Site, browse to the site you have just created. One way to find the app catalog is to choose **site collections** in the left pane, and the app catalog URL will appear in the list of site collections. Once in the app catalog you can add Office add-ins using steps in the next section.

## Publish an add-in to an app catalog

To publish an add-in to an existing app catalog, complete the following steps.

1. Browse to the app catalog:

    - Open the App Catalog Site. If you are not sure how to find it, open the Classic SharePoint admin center page and then choose **site collections** in the left pane. The app catalog URL will be listed along with the other site collections.

2. Choose **Apps for Office** in the left pane.
3. In the **Apps for Office** page, choose **new**.
4. In the **Add a document** dialog, select the **Choose Files** button.
5. Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload.

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

Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy. For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).
