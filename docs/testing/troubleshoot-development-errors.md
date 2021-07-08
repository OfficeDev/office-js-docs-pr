---
title: Troubleshoot development errors with Office Add-ins
description: 'Learn how to troubleshoot development errors in Office Add-ins.'
ms.date: 06/11/2021
localization_priority: Normal
---

# Troubleshoot development errors with Office Add-ins

## Add-in doesn't load in task pane or other issues with the add-in manifest

See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.

## Changes to add-in commands including ribbon buttons and menu items do not take effect

If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer. 

#### For Windows:

Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.

#### For Mac:

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### For iOS:

Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.

## Changes to static files, such as JavaScript, HTML, and CSS do not take effect

The browser may be caching these files. To prevent this, turn off client-side caching when developing. The details will depend on what kind of server you are using. In most cases, it involves adding certain headers to the HTTP Responses. We suggest the following set.

- Cache-Control: "private, no-cache, no-store"
- Pragma: "no-cache"
- Expires: "-1"

For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js). For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

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

Check the reference documentation for the property to see if it is read only. Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only. If you attempt to set a read-only property, the write operation will fail silently, with no error thrown. The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## Getting error: "This add-in is no longer available"

The following are some of the causes of this error. If you discover additional causes, please tell us with the feedback tool at the bottom of the page.

- If you are using Visual Studio, there may be a problem with the sideloading. Close all instances of the Office host and Visual Studio. Restart Visual Studio and try pressing F5 again.
- The add-in's manifest has been removed from its deployment location, such as Centralized Deployment, a SharePoint catalog, or a network share.
- The value of the [ID](../reference/manifest/id.md) element in the manifest has been changed directly in the deployed copy. If for any reason, you want to change this ID, first remove the add-in from the Office host, then replace the original manifest with the changed manifest. You many need to clear the Office cache to remove all traces of the original. See the section [Changes to add-in commands including ribbon buttons and menu items do not take effect](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) earlier in this article.
- The add-in's manifest has a `resid` that is not defined anywhere in the [Resources](../reference/manifest/resources.md) section of the manifest, or there is a mismatch in the spelling of the `resid` between where it is used and where it is defined in the `<Resources>` section.
- There is a `resid` attribute somewhere in the manifest with more than 32 characters. A `resid` attribute, and the `id` attribute of the corresponding resource in the `<Resources>` section, cannot be more than 32 characters.
- The add-in has a custom Add-in Command but you are trying to run it on a platform that doesn't support them. For more information, see [Add-in commands requirement sets](../reference/requirement-sets/add-in-commands-requirement-sets.md).

## Add-in doesn't work on Edge but it works on other browsers

See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).

## Excel add-in throws errors, but not consistently

See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.

## Manifest schema validation errors in Visual Studio projects

If you are using newer features that require changes to the manifest file, you may get validation errors in Visual Studio. For example, when adding the `<Runtimes>` element to implement the shared JavaScript runtime, you may see the following validation error.

**The element 'Host' in namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides' has invalid child element 'Runtimes' in namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides'**

If this occurs, you can update the XSD files that Visual Studio uses to the latest versions. The latest schema versions are at [[MS-OWEMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### Locate the XSD files

1. Open your project in Visual Studio.
1. In **Solution Explorer**, open the manifest.xml file. The manifest is typically in the first project under your solution.
1. Choose **View** > **Properties Window** (F4).
1. In the **Properties Window**, choose the ellipsis (...) to open the **XML Schemas** editor. Here you can find the exact folder location of all schema files your project uses.

### Update the XSD files

1. Open the XSD file you want to update in a text editor. The schema name from the validation error will correlate to the XSD file name. For example, open **TaskPaneAppVersionOverridesV1_0.xsd**.
1. Locate the updated schema at [[MS-OWEMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). For example, TaskPaneAppVersionOverridesV1_0 is at [taskpaneappversionoverrides Schema](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Copy the text into your text editor.
1. Save the updated XSD file.
1. Restart Visual Studio to pick up the new XSD file changes.

You can repeat the previous process for any additional schemas that are out-of-date.

## See also

- [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md)
- [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Debug Office Add-ins on iPad and Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Microsoft Office Add-in Debugger Extension for Visual Studio Code](debug-with-vs-extension.md)
- [Validate an Office Add-in's manifest](troubleshoot-manifest.md)
- [Debug your add-in with runtime logging](runtime-logging.md)
- [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md)
