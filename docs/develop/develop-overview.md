---
title: Develop Office Add-ins
description: An introduction to developing Office Add-ins.
ms.date: 10/22/2019
localization_priority: Priority
---

# Develop Office Add-ins

Office Add-ins can extend the functionality of Office applications and interact with content in Office documents. You can use familiar web technologies to create Office Add-ins that extend and interact with Word, Excel, PowerPoint, OneNote, Project, or Outlook, and your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser. This article provides an introduction to developing Office Add-ins.

> [!TIP]
> If you haven't already done so, please review [Office Add-ins platform overview](../overview/office-add-ins.md) for information that sets context for the topics covered in this article.

## Core development concepts 

...

See: [Components of an Office Add-in](../overview/office-add-ins.md#components-of-an-office-add-in)

### Office Add-ins manifest

An Office Add-in's XML manifest file defines the settings and capabilities of the add-in. You can configure the manifest to specify things such as:

- Metadata that describes the add-in (for example, ID, version, description, display name, default locale)
- The Office applications where the add-in will run
- Permissions that the add-in requires
- How the add-in integrates with Office, including any custom UI that the add-in creates (for example, custom tabs, ribbon buttons)
- Location of images that the add-in uses for branding and command iconography
- Dimensions of the add-in (for example, dimensions for content add-ins, requested height for Outlook add-ins)
- Rules that specify when the add-in activates in the context of a message or appointment (for Outlook add-ins only)

For detailed information about the manifest, see [Office Add-ins XML manifest](add-in-manifests.md).

### Extending the Office UI

Your Office Add-in can extend the Office UI by using add-in commands and HTML containers such as task panes, content add-ins, or dialog boxes.

- [Add-in commands](../design/add-in-commands.md) can be used to add custom tabs, buttons, or menus to the default ribbon in Office, or to extend the default context menu that appears when users right-click text in an Office document or an object in Excel. When users select an add-in command, they initiate the task that the add-in command specifies, such as running JavaScript code, opening a task pane, or launching a dialog box.

- HTML containers like [task panes](../design/task-pane-add-ins.md), [content add-ins](../design/content-add-ins.md), and [dialog boxes](../design/dialog-boxes.md) can be used to display custom UI and expose additional functionality within an Office application. The content and functionality of each task pane, content add-in, or dialog box derives from a web page that you specify. Those web pages can use the Office JavaScript API to interact with content in the Office document where the add-in is running, and can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.

For detailed information about extending the Office UI, see [Design Office Add-ins](../design/add-in-design.md).

### Office JavaScript APIs

An Office Add-in can use the Office JavaScript API to interact with content in the Office document where the add-in is running. 

..
Excel JavaScript API: Introduced with Office 2016, the Excel JavaScript API provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.

Common APIs: Introduced with Office 2013, the Common API can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.
...

The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.  

(How to use these docs - host-specific guidance vs common guidance)


--

Host-specific JavaScript API - Host-specific APIs for Excel and Word provide strongly-typed objects that you can use to access specific elements in the host application. For example, the Excel API contains objects that represent worksheets, ranges, tables, charts, and more.
Common API - Introduced with Office 2013, the Common API enables you to access features such as:
-UI
-Dialogs
-Client settings that are common across multiple types of Office applications

Custom functions use a slightly different programming model and will be covered in a later unit.

### API requirement sets

...

## Creating an Office Add-in 

(Tools & Tech stacks)
...

## Exploring the APIs with Script Lab

...
[Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md)

## Next steps

...

(How to use these docs - host-specific guidance vs common guidance)

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
* [Design Office Add-ins](../design/add-in-design.md)
* [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
* [Publish Office Add-ins](../publish/publish.md)