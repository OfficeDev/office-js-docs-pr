---
title: Publish task pane and content add-ins to a SharePoint app catalog
description: To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the app catalog for their organization.
ms.date: 06/05/2019
localization_priority: Priority
---

# Publish task pane and content add-ins to a SharePoint app catalog

An app catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the app catalog for their organization. When an administrator registers an app catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.

> [!IMPORTANT]
> - App catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.
> - If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.
> - App catalogs on SharePoint are not supported for Office for Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).

## Create an app catalog

Complete the steps in one of the following sections to create an app catalog with on-premises SharePoint Server or on Office 365.

### To create an app catalog for on-premises SharePoint Server

To create the SharePoint app catalog, follow the instructions at [Configure the App Catalog site for a web application](https://docs.microsoft.com/en-us/sharepoint/administration/manage-the-app-catalog).

Once you have created the app catalog follow the steps to [publish an Office Add-in](#publish-an-office-add-in).

### To create an app catalog on Office 365

1. Go to the Microsoft 365 admin center. For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).

2. On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.

    > [!NOTE]
    > You need to use the Classic SharePoint admin center to create the catalog. If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.

3. In the left task pane, choose  **apps**.

4. On the **apps** page, choose **App Catalog**.
    > [!NOTE]
    > If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.

5. On the **App Catalog Site** page, choose **OK** to accept the default option and create a new app catalog site.

6. On the **Create App Catalog Site Collection** page, specify the title of your App Catalog site.

7. Specify the **Web Site Address**.

8. Specify an **Administrator**.

9. Set the **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your app catalog site.)

10. Choose **OK**.

## Publish an Office Add-in

Complete the steps in one of the following sections to publish an Office Add-in to an app catalog on Office 365 or on-premises SharePoint Server.

### To publish an Office add-in to a SharePoint app catalog on Office 365

1. Go to the Microsoft 365 admin center. For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).
2. On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.
    > [!NOTE]
    > You need to use the Classic SharePoint admin center to create the catalog. If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.
3. In the left task pane, choose  **apps**.
4. On the **apps** page, choose **App Catalog**.
5. Choose **Distribute apps for Office**.
6. In the **Apps for Office** page, choose **New**.
7. In the **Add a document** dialog, select the **Choose Files** button.
8. Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.
9. In the **Add a document** dialog, choose **OK**.

### To publish an add-in to an app catalog with on-premises SharePoint Server

1. Open the **Central Administration** page.
2. In the left task pane, choose **Apps**.
3. On the **Apps** page, under **App Management**, choose **Manage App Catalog**.
4. On the **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application** Selector.
5. Choose the URL under the **Site URL** to open the app catalog site.
6. Choose **Distribute apps for Office**.
7. In the **Apps for Office** page, choose **New**.
8. In the **Add a document** dialog, select the **Choose Files** button.
9. Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.
10. In the **Add a document** dialog, choose **OK**.

## Insert Office Add-ins from the app catalog

For online Office applications, you can find Office Add-ins from the app catalog by completing the following steps.

1. Open the online Office application (Excel, PowerPoint, or Word).
2. Create or open a document.
3. Choose **Insert** > **Add-ins**.
4. In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.
    The Office Add-ins are listed.
5. Choose an Office Add-in and then choose **Add**.

For Office applications on the desktop, you can find Office Add-ins from the app catalog by completing the following steps.

1. Open the desktop Office application (Excel, Word, or PowerPoint)
2. Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.
3. Enter the URL of the SharePoint app catalog in the **Catalog Url** box and choose **Add catalog**.
    Use the shorter form of the URL. For example, if the URL of the SharePoint app catalog is:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    Specify just the URL of the parent site collection:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. Close and reopen the Office application. 
5. Choose **Insert** > **Get Add-ins**.
4. In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.
    The Office Add-ins are listed.
5. Choose an Office Add-in and then choose **Add**.

Alternatively, an administrator can specify an app catalog on SharePoint by using group policy. For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).
