---
title: Troubleshoot development errors with Office Add-ins
description: Learn how to troubleshoot development errors in Office Add-ins.
ms.topic: troubleshooting-problem-resolution
ms.date: 02/12/2025
ms.localizationpriority: medium
---

# Troubleshoot development errors with Office Add-ins

Here's a list of common issues you may encounter while developing an Office Add-in.

> [!TIP]
> Clearing the Office cache often fixes issues related to stale code. This guarantees the latest manifest is uploaded, using the current file names, menu text, and other command elements. To learn more, see [Clear the Office cache](clear-cache.md).

## Add-in doesn't load in task pane or other issues with the add-in manifest

See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.

## Ribbon customizations are not rendering as expected

- With the add-in sideloaded and running, paste the URLs for the add-in's ribbon icons into a browser's navigation bar and see if the icon files open.
- By default, add-in errors connected to the Office UI are suppressed. You can turn on these error messages with the following steps.

   1. With the add-in removed, open the **File** tab of the Office application.
   1. Select **Options**.
   1. In the **Options** dialog, select **Advanced**.
   1. In the **General** section (the **Developers** section for Outlook), enable **Show add-in user interface errors**.

   Sideload the add-in again and see if there are any errors.

## Changes to add-in commands including ribbon buttons and menu items do not take effect

Clearing the cache helps ensure the latest version of your add-in's manifest is being used. To clear the Office cache, follow the instructions in [Clear the Office cache](clear-cache.md). If you're using Office on the web, clear your browser's cache through the browser's UI.

## Add-in commands from old development add-ins stay on ribbon even after the cache is cleared

Sometimes buttons or menus from an add-in that you were developing in the past appears on the ribbon when you run an Office application even after you have cleared the cache. Try these techniques:

- If you develop add-ins on more than one computer and your user settings are synchronized across the computers, try [clearing the Office cache](clear-cache.md) on all the computers. Shut down all Office applications on all the computers, and then clear the cache on all of them before you open any Office application on any of them.
- If you [published the manifest of the old add-in to a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md), shut down all Office applications, clear the cache, and then *be sure that the manifest for the add-in is removed from the shared folder*.  

## Changes to static files, such as JavaScript, HTML, and CSS do not take effect

The browser may be caching these files. To prevent this, turn off client-side caching when developing. The details will depend on what kind of server you are using. In most cases, it involves adding certain headers to the HTTP Responses. We suggest the following set.

- Cache-Control: "private, no-cache, no-store"
- Pragma: "no-cache"
- Expires: "-1"

For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js). For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNETCore-WebAPI/Views/Shared/_Layout.cshtml).

If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

If these steps don't seem to work at first, you may need to clear the browser's cache. Do this through the UI of the browser. Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI. If that happens, run the following command in a Windows Command Prompt.

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## Changes made to property values don't happen and there is no error message

Check the reference documentation for the property to see if it is read-only. Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only. If you attempt to set a read-only property, the write operation will fail silently, with no error thrown. The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## Getting error: "This add-in is no longer available"

The following are some of the causes of this error. If you discover additional causes, please tell us with the feedback tool at the bottom of the page.

- If you're using Visual Studio, there may be a problem with the sideloading. Close all instances of the Office host and Visual Studio. Restart Visual Studio and try pressing <kbd>F5</kbd> again.
- The add-in's manifest has been removed from its deployment location, such as Centralized Deployment, a SharePoint catalog, or a network share.
- If the add-in only manifest is being used, one of the following may apply.

  - The value of the [ID](/javascript/api/manifest/id) element in the manifest has been changed directly in the deployed copy. If for any reason, you want to change this ID, first remove the add-in from the Office host, then replace the original manifest with the changed manifest. You many need to clear the Office cache to remove all traces of the original. See the [Clear the Office cache](clear-cache.md) article for instructions on clearing the cache for your operating system.
  - The add-in's manifest has a `resid` that isn't defined anywhere in the [Resources](/javascript/api/manifest/resources) section of the manifest, or there is a mismatch in the spelling of the `resid` between where it is used and where it is defined in the **\<Resources\>** section.
  - There is a `resid` attribute somewhere in the manifest with more than 32 characters. A `resid` attribute, and the `id` attribute of the corresponding resource in the **\<Resources\>** section, cannot be more than 32 characters.

- The add-in has a custom Add-in Command but you are trying to run it on a platform that doesn't support them. For more information, see [Add-in commands requirement sets](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets).

## Add-in doesn't work on Edge but it works on other browsers

See [Troubleshoot EdgeHTML and WebView2 (Microsoft Edge) issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshoot-edgehtml-and-webview2-microsoft-edge-issues).

## Excel add-in throws errors, but not consistently

See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.

## Word add-in throws errors or displays broken behavior

See [Troubleshoot Word add-ins](../word/word-add-ins-troubleshooting.md) for possible causes.

## Add-in only manifest schema validation errors in Visual Studio projects

If you're using newer features that require changes to the add-in only manifest file, you may get validation errors in Visual Studio. For example, when adding the **\<Runtimes\>** element to implement the [shared runtime](runtimes.md#shared-runtime), you may see the following validation error.

**The element 'Host' in namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides' has invalid child element 'Runtimes' in namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides'**

If this occurs, you can update the XSD files that Visual Studio uses to the latest versions. The latest schema versions are at [[MS-OWEMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### Locate the XSD files

1. Open your project in Visual Studio.
1. In **Solution Explorer**, open the manifest.xml file. The manifest is typically in the first project under your solution.
1. Select **View** > **Properties Window** (<kbd>F4</kbd>).
1. Set the cursor selection in the manifest.xml so that the **Properties** window shows the **XML Document** properties.
1. In the **Properties** window, select the **Schemas** property, then select the ellipsis (...) to open the **XML Schemas** editor. Here you can find the exact folder location of all schema files your project uses.

:::image type="content" source="../images/visual-studio-xml-document-properties.png" alt-text="Properties window showing the XML document properties.":::

### Update the XSD files

1. Open the XSD file you want to update in a text editor. The schema name from the validation error will correlate to the XSD file name. For example, open **TaskPaneAppVersionOverridesV1_0.xsd**.
1. Locate the updated schema at [[MS-OWEMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). For example, TaskPaneAppVersionOverridesV1_0 is at [taskpaneappversionoverrides Schema](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Copy the text into your text editor.
1. Save the updated XSD file.
1. Restart Visual Studio to pick up the new XSD file changes.

You can repeat the previous process for any additional schemas that are out-of-date.

## When working offline, no Office APIs work

When you're loading the Office JavaScript Library from a local copy instead of from the CDN, the APIs may stop working if the library isn't up-to-date. If you have been away from a project for a while, reinstall the library to get the latest version. The process varies according to your IDE. Choose one of the following options based on your environment.

- **Visual Studio**: Follow these steps to update the NuGet package.
    1. Choose **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.
    1. Choose the **Updates** tab.
    1. Select "Microsoft.Office.js". Ensure the package source is from nuget.org.
    1. In the left pane, choose **Install** and complete the package update process.
- **Any other IDE**: Get the latest npm packages [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) and [@types/office-js](https://www.npmjs.com/package/@types/office-js).

## See also

- [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md)
- [Sideload an Office Add-in on Mac](sideload-an-office-add-in-on-mac.md)  
- [Sideload an Office Add-in on iPad](sideload-an-office-add-in-on-ipad.md)  
- [Debug Office Add-ins on a Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Validate an Office Add-in's manifest](troubleshoot-manifest.md)
- [Debug your add-in with runtime logging](runtime-logging.md)
- [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md)
- [Runtimes in Office Add-ins](runtimes.md)
- [Microsoft Q&A (Office Development)](https://aka.ms/office-addins-dev-questions)
