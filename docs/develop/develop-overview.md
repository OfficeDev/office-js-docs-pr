---
title: Develop Office Add-ins
description: An introduction to developing Office Add-ins.
ms.date: 10/22/2019
localization_priority: Priority
---

# Develop Office Add-ins

Office Add-ins can extend the UI and functionality of Office applications and interact with content in Office documents. You can use familiar web technologies to create Office Add-ins that extend and interact with Word, Excel, PowerPoint, OneNote, Project, or Outlook, and your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser. This article provides an introduction to developing Office Add-ins.

> [!TIP]
> If you haven't already done so, please review [Office Add-ins platform overview](../overview/office-add-ins.md) for information that sets context for the topics covered in this article.

## Core development concepts 

As described in [Office Add-ins platform overview](../overview/office-add-ins.md), an Office Add-in consists of two parts:

- The add-in manifest (an XML file) that defines the settings and capabilities of the add-in.

- The web application that defines the UI and functionality of add-in components such as task panes, content add-ins, and dialog boxes.

The web application can use the Office JavaScript API to interact with content in the Office document where the add-in is running, and can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.

### Defining an add-in's settings and capabilities

An Office Add-in's manifest (an XML file) defines the settings and capabilities of the add-in. You can configure the manifest to specify things such as:

- Metadata that describes the add-in (for example, ID, version, description, display name, default locale)
- The Office applications where the add-in will run
- Permissions that the add-in requires
- How the add-in integrates with Office, including any custom UI that the add-in creates (for example, custom tabs, ribbon buttons)
- Location of images that the add-in uses for branding and command iconography
- Dimensions of the add-in (for example, dimensions for content add-ins, requested height for Outlook add-ins)
- Rules that specify when the add-in activates in the context of a message or appointment (for Outlook add-ins only)

For detailed information about the manifest, see [Office Add-ins XML manifest](add-in-manifests.md).

### Extending the Office UI

An Office Add-in can extend the Office UI by using add-in commands and HTML containers such as task panes, content add-ins, or dialog boxes.

- [Add-in commands](../design/add-in-commands.md) can be used to add custom tabs, buttons, or menus to the default ribbon in Office, or to extend the default context menu that appears when users right-click text in an Office document or an object in Excel. When users select an add-in command, they initiate the task that the add-in command specifies, such as running JavaScript code, opening a task pane, or launching a dialog box.

- HTML containers like [task panes](../design/task-pane-add-ins.md), [content add-ins](../design/content-add-ins.md), and [dialog boxes](../design/dialog-boxes.md) can be used to display custom UI and expose additional functionality within an Office application. The content and functionality of each task pane, content add-in, or dialog box derives from a web page that you specify. Those web pages can use the Office JavaScript API to interact with content in the Office document where the add-in is running, and can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.

For detailed information about extending the Office UI, see [Design Office Add-ins](../design/add-in-design.md).

### Interacting with content in an Office document

An Office Add-in can use the Office JavaScript APIs to interact with content in the Office document where the add-in is running. 

#### Accessing the Office JavaScript API library

The CDN for the Office JavaScript API library resides at `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

#### API object models

The Office JavaScript APIs include two distinct object models:

- **Host-specific** APIs (introduced with Office 2016) provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application. For example, you can use the Excel JavaScript APIs to access worksheets, ranges, tables, charts, and more. Host-specific APIs are currently available for [Excel](../reference/overview/excel-add-ins-reference-overview.md), [Word](../reference/overview/word-add-ins-reference-overview.md), [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md), and [PowerPoint](..//reference/overview/powerpoint-add-ins-reference-overview.md). This object model uses [promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).

- **Common** APIs (introduced with Office 2013) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications. This object model uses [callbacks](https://developer.mozilla.org/en-US/docs/Glossary/Callback_function). For details about the Common APIs, which include APIs for interacting with Outlook, see [Office JavaScript API object model](office-javascript-api-object-model.md)

> [!NOTE]
> Please note the following:
> 
> - [Outlook APIs](/outlook/add-ins/apis) are accessed by using the Common API syntax.
> 
> - Excel Custom functions run within a unique runtime that prioritizes execution of calculations, and therefore uses a slightly different programming model. For details, see [Custom functions architecture](../excel/custom-functions-architecture.md).

For additional information about the Office JavaScript APIs, see [Understanding the JavaScript API for Office](understanding-the-javascript-api-for-office.md).

#### API requirement sets

...

## Creating an Office Add-in 

(Tools & Tech stacks)
...

## Exploring APIs with Script Lab

...
[Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md)

## Next steps

...

> [!TIP]
> (How to use these docs - host-specific guidance vs common guidance)

Explore content in this section of the docs to learn more about scenarios that apply to building any type of add-in

* ...
* ...

Explore content in the host-specific section of the docs for the type of add-in you're building.

* [Excel add-ins documentation](../excel/index.md)
* [OneNote add-ins documentation](../onenote/index.md)
* [Outlook add-ins documentation](../outlook/index.md)
* [PowerPoint add-ins documentation](../powerpoint/index.md)
* [Project add-ins documentation](../project/index.md)
* [Visio add-ins documentation](../visio/index.md)
* [Word add-ins documentation](../word/index.md)

Complete a quick start | Complete a tutorial

Learn more about [testing and debugging Office Add-ins](../testing/test-debug-office-add-ins.md) and [Publishing Office Add-ins](../publish/publish.md).

...

## See also

* [Office Add-ins platform overview](../overview/office-add-ins.md)
* [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md)
* [Understanding the JavaScript API for Office](understanding-the-javascript-api-for-office.md)
* [Design Office Add-ins](../design/add-in-design.md)
* [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
* [Publish Office Add-ins](../publish/publish.md)