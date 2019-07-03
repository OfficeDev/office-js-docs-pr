---
title: JavaScript API for Office
description: ''
ms.date: 05/13/2019
localization_priority: Priority
---

# JavaScript API for Office

The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library, which is a script loader. The office.js library loads the object models that are applicable to the Office application that is running the add-in. You can use the following JavaScript object models:

- **Common APIs** - APIs that were introduced with **Office 2013**. This is loaded for **all Office host applications** and connects your add-in application with the Office client application. The object model contains APIs that are specific to Office clients, and APIs that are applicable to multiple Office client host applications. All of this content is under **Common API**. This object model uses callbacks. 

  **Outlook** also uses the Common API syntax. Everything under the alias Office contains objects you can use to write scripts that interact with content in Office documents, worksheets, presentations, mail items, and projects from your Office Add-ins. You must use these Common APIs if your add-in will target Office 2013 and later. This object model uses callbacks.

- **Host-specific APIs** - APIs that were introduced with **Office 2016**. This object model provides host-specific strongly-typed objects that correspond to familiar objects that you see when you use Office clients, and represents the future of Office JavaScript APIs. The host-specific APIs currently include the the Excel JavaScript API, the OneNote JavaScript API, the PowerPoint JavaScript API, and the Word JavaScript API.

## Supported host applications

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [PowerPoint](overview/powerpoint-add-ins-reference-overview.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [Common API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [Project](overview/project-add-ins-reference-overview.md) supports add-ins made with the JavaScript API, but there's currently no JavaScript API designed specifically for interacting with Project. You can use the Common API to interact with objects and data in Project.

Learn more about [supported hosts and other requirements](../concepts/requirements-for-running-office-add-ins.md).

## Open API specifications

As we design and develop new APIs for Office Add-ins, we'll make them available for your feedback on our [Open API specifications](openspec/openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.

## See also

- [Office JavaScript API reference](/javascript/api/overview/office)
