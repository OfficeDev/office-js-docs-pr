
# JavaScript API for Office

The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library which is a script loader. office.js loads the object models that are applicable to the Office application that is running the add-in. There are two potential JavaScript object object model forms you may use:

1) Common - APIs that were introduced with Office 2013. This is loaded for **all Office host applications** and must be used as it connects your add-in application with the Office client application. The object model contains APIs that are specific to Office clients, and many APIs that are applicable to many Office client host applications. All of the content under [shared](../reference/shared/shared-api.md) and **outlook** are considered the common APIs. The  **Microsoft.Office.WebExtension** namespace (which by default is referenced using the alias [Office](../reference/shared/office.md) in code) contains objects you can use to write script that interacts with content in Office documents, worksheets, presentations, mail items, and projects from your Office Add-ins. You must use these common APIs if your add-in will target Office 2013 and later. This object model form uses callbacks.

2) Host specific - APIs that were introduced with Office 2016. The new object model that provides host specific strongly-typed objects that correspond to familiar objects that you see when you use Office clients. This is the future of Office JavaScript APIs and should be used moving forward. This is currently applicable to [Word](../reference/word/word-add-ins-javascript-reference.md) and **Excel**.  This object model form uses promises.

Select the Office client from the dropdown above the TOC to filter the content based on your target host application.

## Supported host applications
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word

Learn more about [supported hosts and other requirements](../docs/overview/requirements-for-running-office-add-ins.md).
