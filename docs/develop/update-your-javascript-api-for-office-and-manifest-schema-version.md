---
title: Update to the latest JavaScript API for Office library and version 1.1 add-in manifest schema
description: Update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.
ms.date: 12/04/2017
---

# Update to the latest JavaScript API for Office library and version 1.1 add-in manifest schema

This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.

## Use the most up-to-date project files

If you use Visual Studio to develop your add-in, to use the [newest API members](https://docs.microsoft.com/javascript/office/what's-changed-in-the-javascript-api-for-office?view=office-js) of the JavaScript API for Office and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download and install the [Visual Studio 2015 and the latest Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs).

If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.

To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Office 365, SharePoint Online, and Exchange Online.

To download Office, SharePoint, and Exchange SP1 products, see the following:

- [List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products](http://support.microsoft.com/kb/2850036)
    
- [List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products](http://support.microsoft.com/kb/2850035)
    
- [Description of Exchange Server 2013 Service Pack 1](http://support.microsoft.com/kb/2926248)
    

## Updating an Office Add-in project created with Visual Studio

For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you can update a project's files using the  **NuGet Package Manager**, and then update your add-in's HTML pages to reference them. 

Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.


### Update the JavaScript API for Office library files in your project to the newest release


1. In Visual Studio 2015, open or create a new  **Office Add-in** project.
    
      - In the left pane, choose **Update** and complete the package update process.
    
      - Go to step 6.
    
2. Choose  **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.
    
3. In the  **NuGet Package Manager**, select  **nuget.org** for **Package source** and **Upgrade available** for **Filter**. and select Microsoft.Office.js.
    
4. In the left pane, choose **Update** and complete the package update process.
    
5. In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.   


### Update the manifest file in your project to use schema version 1.1

In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
	xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js) and [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).

## Updating an Office Add-in project created with a text editor or other IDE

For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1. 

The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.

You don't need local copies of the JavaScript API for Office files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](http://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.

> [!NOTE] 
> To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).


### Update the JavaScript API for Office library files in your project to use the newest release

1. Open the HTML pages for your add-in in your text editor or IDE.
    
2. In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.   

### Update the manifest file in your project to use schema version 1.1

In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
	xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js) and [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).
    

## See also

- [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md) 
- [Understanding the JavaScript API for Office](understanding-the-javascript-api-for-office.md)    
- [JavaScript API for Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)   
- [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md)
    
