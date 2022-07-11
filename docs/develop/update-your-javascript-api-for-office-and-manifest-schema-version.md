---
title: Update to the latest Office JavaScript API library and version 1.1 add-in manifest schema
description: Update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.
ms.date: 01/14/2022
ms.localizationpriority: medium
---

# Update to the latest Office JavaScript API library and version 1.1 add-in manifest schema

This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.

> [!NOTE]
> Projects created in Visual Studio 2019 will already use version 1.1. However there are occasional minor updates to version 1.1 that you can apply by using the techniques in this article.

## Use the most up-to-date project files

If you use Visual Studio to develop your add-in, to use the newest API members of the Office JavaScript API and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download Visual Studio 2019. To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/). During installation you'll need to select the Office/SharePoint development workload.

If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the content delivery network (CDN) for Office.js and the version of schema referenced in your add-in's manifest.

To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Microsoft 365, SharePoint Online, and Exchange Online.

To download Office, SharePoint, and Exchange SP1 products, see the following:

- [List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products](https://support.microsoft.com/kb/2850036)

- [List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products](https://support.microsoft.com/kb/2850035)

- [Description of Exchange Server 2013 Service Pack 1](https://support.microsoft.com/kb/2926248)

## Updating an Office Add-in project created with Visual Studio

For projects created before the release of v1.1 of the Office JavaScript API and add-in manifest schema, you can update a project's files using the **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.

Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.

### Update the Office JavaScript API library files in your project to the newest release

The following steps will update your Office.js library files to the latest version. The steps use Visual Studio 2019, but they are similar for previous versions of Visual Studio.

1. In Visual Studio 2019, open or create a new **Office Add-in** project.
2. Choose **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.
3. Choose the **Updates** tab.
4. Select Microsoft.Office.js. Ensure the package source is from **nuget.org**.
5. In the left pane, choose **Install** and complete the package update process.

You'll need to take a few additional steps to complete the update. In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API library as follows:

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE]
   > The `/1/` in the `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.

### Update the manifest file in your project to use schema version 1.1

In your add-in's manifest file, update the **xmlns** attribute of the **\<OfficeApp\>** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> After updating the version of the add-in manifest schema to 1.1, you will need to remove the **Capabilities** and **Capability** elements, and replace them with either the [Hosts](/javascript/api/manifest/hosts) and [Host](/javascript/api/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).

## Updating an Office Add-in project created with a text editor or other IDE

For projects created before the release of v1.1 of the Office JavaScript API and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.

The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.

You don't need local copies of the Office JavaScript API files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](https://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.

> [!NOTE]
> To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).

### Update the Office JavaScript API library files in your project to use the newest release

1. Open the HTML pages for your add-in in your text editor or IDE.

2. In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API library as follows:

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.

### Update the manifest file in your project to use schema version 1.1

In your add-in's manifest file, update the **xmlns** attribute of the **\<OfficeApp\>** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> After updating the version of the add-in manifest schema to 1.1, you will need to remove the **Capabilities** and **Capability** elements, and replace them with either the [Hosts](/javascript/api/manifest/hosts) and [Host](/javascript/api/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).

## See also

- [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md) ]
- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
- [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md)
