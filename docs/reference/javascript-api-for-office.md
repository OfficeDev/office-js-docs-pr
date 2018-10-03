# JavaScript API for Office

The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library, which is a script loader. The office.js library loads the object models that are applicable to the Office application that is running the add-in. You can use the following JavaScript object models:

- **Common APIs** - APIs that were introduced with **Office 2013**. This is loaded for **all Office host applications** and connects your add-in application with the Office client application. The object model contains APIs that are specific to Office clients, and APIs that are applicable to multiple Office client host applications. All of this content is under **Shared API**. 

  **Outlook** also uses the common API syntax. Everything under the alias Office contains objects you can use to write scripts that interact with content in Office documents, worksheets, presentations, mail items, and projects from your Office Add-ins. You must use these common APIs if your add-in will target Office 2013 and later. This object model uses callbacks.

- **Host-specific APIs** - APIs that were introduced with **Office 2016**. This object model provides host-specific strongly-typed objects that correspond to familiar objects that you see when you use Office clients, and represents the future of Office JavaScript APIs. The host-specific APIs currently include the Word JavaScript API and the Excel JavaScript API.

## Supported host applications

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [Shared API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint and Project](requirement-sets/powerpoint-and-project-note.md) support add-ins made with the JavaScript API. However, they currently do not have host-specific APIs. You interact with these hosts through the Shared API.

Learn more about [supported hosts and other requirements](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## Open API specifications

As we design and develop new APIs for Office Add-ins, we'll make them available for your feedback on our [Open API specifications](openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.

## See also

- [Office JavaScript API reference](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)