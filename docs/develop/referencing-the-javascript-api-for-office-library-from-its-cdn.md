
# Referencing the JavaScript API for Office library from its content delivery network (CDN)


The [JavaScript API for Office](../../reference/javascript-api-for-office.md) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. 


The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

The  `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). 

If you plan on publishing your Office Add-in from the Office Store, you must use this CDN reference. Local references are only appropriate for internal, development and debugging scenarios.

> **Important:** 
When developing an  Add-in for any Office host application, it is important to reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures the API is fully initialized prior to any body elements. Office hosts require that Add-Ins initialize within 5 seconds of activation. Crossing this threshold results in the Add-In being declared unresponsive and an error message displayed to the user.       

## Additional resources



- [Understanding the JavaScript API for Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Office Add-ins platform overview](../../docs/overview/office-add-ins.md)
    
- [Office Add-ins development lifecycle](../../docs/design/add-in-development-lifecycle.md)
    
- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
    
