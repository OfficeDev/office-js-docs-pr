
# Referencing the JavaScript API for Office library from its content delivery network (CDN)


The [JavaScript API for Office](../../reference/javascript-api-for-office.md) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. When developing an Office Add-in for any Office host application, you should reference the JavaScript API for Office library inside the `<head>` tag of the web page (such as an .html, .aspx, or .php file) that implements the UI of your add-in. To do that, add a `script` tag with its `src` attribute set to the following CDN URL.



```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

The  `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files] (../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).
The first time your add-in loads, the JavaScript API for Office library files will be downloaded and cached to make sure that your add-in is using the most up-to-date implementation of Office.js and application-specific .js files.
The default Home.html file in your project will contain the appropriate  `script` tag if you develop your add-in with the **Add-in for Office** project template files provided with the latest Visual Studio with the [latest Microsoft Office Developer Tools update](https://www.visualstudio.com/features/office-tools-vs).

## Additional resources



- [Understanding the JavaScript API for Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Office Add-ins platform overview](../../docs/overview/office-add-ins.md)
    
- [Office Add-ins development lifecycle](../../docs/design/add-in-development-lifecycle.md)
    
- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
    
