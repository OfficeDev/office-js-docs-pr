---
title: Publish task pane and content add-ins to a SharePoint app catalog
description: To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the app catalog for their organization.
ms.date: 11/08/2021
ms.localizationpriority: medium
---

# Publish task pane and content add-ins to a SharePoint app catalog

An app catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the app catalog for their organization. When an administrator registers an app catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.

> [!IMPORTANT]
>
> - App catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [Office add-in XML manifest](../develop/xml-manifest-overview.md), such as add-in commands.
> - If you’re targeting a cloud or hybrid environment, we recommend that you [use Integrated Apps via the Microsoft 365 admin center](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) to publish your add-ins.
> - App catalogs on SharePoint are not supported in Office on Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).

## Create an app catalog

Complete the steps in one of the following sections to create an app catalog with on-premises SharePoint Server or on Microsoft 365.

### To create an app catalog for on-premises SharePoint Server

To create the SharePoint app catalog, follow the instructions at [Configure the App Catalog site for a web application](/sharepoint/administration/manage-the-app-catalog).

Once you have created the app catalog follow the steps to [publish an Office Add-in](#publish-an-office-add-in).

### To create an app catalog on Microsoft 365

To create the SharePoint app catalog, follow the instructions at [Create the App Catalog site collection](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection). Once you have created the app catalog, follow the steps in the next section to publish an Office Add-in.

## Publish an Office Add-in

Complete the steps in one of the following sections to publish an Office Add-in to an app catalog on Microsoft 365 or on-premises SharePoint Server.

### To publish an Office Add-in to a SharePoint app catalog on Microsoft 365

1. Go to the [Active sites page of the new SharePoint admin center](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true) and sign in with an account that has [admin permissions](/sharepoint/sharepoint-admin-role) for your organization.

    > [!NOTE]
    > If you have Microsoft 365 operated by 21Vianet (China), [sign in to the Microsoft 365 admin center](https://go.microsoft.com/fwlink/p/?linkid=850627), then browse to the SharePoint admin center and open the More features page.

1. Open the app catalog site by selecting its URL in the URL column.

    > [!NOTE]
    > If you just created the app catalog site in the previous section, it can take a few minutes for the site to finish setting up.

1. Choose **Distribute apps for Office**.
1. In the **Apps for Office** page, choose **New**.
1. In the **Add a document** dialog, select the **Choose Files** button.
1. Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.
1. In the **Add a document** dialog, choose **OK**.

### To publish an add-in to an app catalog with on-premises SharePoint Server

1. Open the **Central Administration** page.
1. In the left task pane, choose **Apps**.
1. On the **Apps** page, under **App Management**, choose **Manage App Catalog**.
1. On the **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application** Selector.
1. Choose the URL under the **Site URL** to open the app catalog site.
1. Choose **Distribute apps for Office**.
1. In the **Apps for Office** page, choose **New**.
1. In the **Add a document** dialog, select the **Choose Files** button.
1. Locate and specify the [XML manifest](../develop/xml-manifest-overview.md) file to upload and choose **Open**.
1. In the **Add a document** dialog, choose **OK**.

## Insert Office Add-ins from the app catalog

For online Office applications, you can find Office Add-ins from the app catalog by completing the following steps.

1. Open the online Office application (Excel, PowerPoint, or Word).
1. Create or open a document.
1. Choose **Insert** > **Add-ins**.
1. In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.
    The Office Add-ins are listed.
1. Choose an Office Add-in and then choose **Add**.

For Office applications on the desktop, you can find Office Add-ins from the app catalog by completing the following steps.

1. Open the desktop Office application (Excel, Word, or PowerPoint)
1. Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.
1. Enter the URL of the SharePoint app catalog in the **Catalog Url** box and choose **Add catalog**.
    Use the shorter form of the URL. For example, if the URL of the SharePoint app catalog is:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`

    Specify just the URL of the parent site collection:
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
1. Close and reopen the Office application.
1. Choose **Insert** > **Get Add-ins**.
1. In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.
    The Office Add-ins are listed.
1. Choose an Office Add-in and then choose **Add**.

Alternatively, an administrator can specify an app catalog on SharePoint by using Group Policy. The relevant policy settings are available in the [Administrative Template files (ADMX/ADML) for Microsoft 365 Apps, Office LTSC 2021, Office 2019, and Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) and be found under **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs**.
