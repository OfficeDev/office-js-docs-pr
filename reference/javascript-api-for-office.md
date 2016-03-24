
# JavaScript API for Office reference

The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library, which is a script loader. The office.js  library loads the object models that are applicable to the Office application that is running the add-in. You can use these JavaScript object models:


1. Common (required) - APIs that were introduced with Office 2013. This is loaded for **all Office host applications** and connects your add-in application with the Office client application. The object model contains APIs that are specific to Office clients, and APIs that are applicable to multiple Office client host applications. All the content under [shared](../reference/shared/shared-api.md) and **outlook** are considered the common APIs. The  **Microsoft.Office.WebExtension** namespace (which by default is referenced using the alias [Office](../reference/shared/office.md) in code) contains objects you can use to write script that interacts with content in Office documents, worksheets, presentations, mail items, and projects from your Office Add-ins. You must use these common APIs if your add-in will target Office 2013 and later. This object model uses callbacks.

1. Host-specific - APIs that were introduced with Office 2016. This object model provides host-specific strongly-typed objects that correspond to familiar objects that you see when you use Office clients. This model represents the future of Office JavaScript APIs and will apply moving forward. This is currently applicable to **Word** and **Excel**. This object model uses promises.

Select the Office client from the drop-down above the TOC to filter the content based on your target host application.

## Supported host applications
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word

Learn more about [supported hosts and other requirements](../docs/overview/requirements-for-running-office-add-ins.md).
