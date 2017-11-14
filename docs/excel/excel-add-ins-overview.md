---
title: Excel add-ins overview
description: 
ms.date: 11/20/2017 
---


# Excel add-ins overview

An Excel add-in allows you to extend Excel application functionality across multiple platforms including Office for Windows, Office Online, Office for the Mac, and Office for the iPad. Use Excel add-ins within a workbook to:

- Interact with Excel objects, read and write Excel data. 
- Extend functionality using web based task pane or content pane 
- Add custom ribbon buttons or contextual menu items
- Provide richer interaction using dialog window 

The Office Add-ins platform provides the framework and Office.js JavaScript APIs that enable you to create and run Excel add-ins. By using the Office Add-ins platform to create your Excel add-in, you'll get the following benefits:

* **Cross-platform support**: Excel add-ins run in Office for Windows, Mac, iOS, and Office Online.
* **Centralized deployment**: Admins can quickly and easily deploy Excel add-ins to users throughout an organization.
* **Single sign on (SSO)**: Easily integrate your Excel add-in with the Microsoft Graph.
* **Use of standard web technology**: Create your Excel add-in using familiar web technologies such as HTML, CSS, and JavaScript.
* **Distribution via the Office Store**: Share your Excel add-in with a broad audience by publishing it to the [Office Store](https://store.office.com/en-us/appshome.aspx).

> [!NOTE]
> Excel add-ins are different from COM and VSTO add-ins, which are earlier Office integration solutions that run only on Office for Windows. Unlike COM add-ins, Excel add-ins do not require you to install any code on a user's device, or within Excel. 

## Components of an Excel add-in 

An Excel add-in includes two basic components: a web application and a configuration file, called a manifest file. 

The web application uses the [JavaScript API for Office](../../reference/javascript-api-for-office.md) to interact with objects in Excel, and can also facilitate interaction with online resources. For example, an add-in can perform any of the following tasks:

* Create, read, update, and delete data in the workbook (worksheets, ranges, tables, charts, named items, and more).
* Perform user authorization with an online service by using the standard OAuth 2.0 flow.
* Issue API requests to Microsoft Graph or any other API.

The web application can be hosted on any web server, and can be built using client-side frameworks (such as Angular, React, jQuery) or server-side technologies (such as ASP.NET, Node.js, PHP).

The [manifest](../overview/add-in-manifests.md) is an XML configuration file that defines how the add-in integrates with Office clients by specifying settings and capabilities such as: 

* The URL of the add-in's web application.
* The add-in's display name, description, ID, version, and default locale.
* How the add-in integrates with Excel, including any custom UI that the add-in creates (ribbon buttons, context menus, and so on).
* Permissions that the add-in requires, such as reading and writing to the document.

To enable end-users to install and use an Excel add-in, you must publish its manifest to either the Office Store or to an Add-ins catalog. 

## Capabilities of an Excel add-in

In addition to interacting with the content in the workbook, Excel add-ins can add custom ribbon buttons or menu commands, insert task panes, open dialog boxes, and even embed rich, web-based objects such as charts or interactive visualizations within a worksheet, as shown in the following screenshots. For more information about each of these capabilities, see [Extend Excel functionality](excel-add-ins-extend-excel.md).

**Custom ribbon buttons**

![Add-in commands](../images/excel-add-in-commands-script-lab.png)

**Task pane**

![Add-in task pane](../images/excel-add-in-task-pane-insights.png)

**Dialog box**

![Add-in dialog box](../images/excel-add-in-dialog-choose-number.png)

**Content add-in**

![Content add-in](../images/excel-add-in-content-map.png)

## JavaScript APIs to interact with workbook content

An Excel add-in interacts with objects in Excel by using the [JavaScript API for Office](../../reference/javascript-api-for-office.md), which includes two JavaScript object models:

* **Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../../reference/excel/excel-add-ins-reference-overview.md) provides strongly-typed Excel objects that you can use to access worksheets, ranges, tables, charts, and more. 

* **Shared API**: Introduced with Office 2013, the shared API enables you to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint. Because the shared API does provide limited functionality for Excel interaction, you can use it if your add-in needs to run on Excel 2013.

## Next steps

Get started by [creating your first Excel add-in](excel-add-ins-get-started-overview.md). Then, learn about the [core concepts](excel-add-ins-core-concepts.md) of building Excel add-ins.

## Additional resources

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Best practices for developing Office Add-ins](../overview/add-in-development-best-practices.md)
- [Design guidelines for Office Add-ins](../design/add-in-design.md)
- [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md)
- [Excel JavaScript API reference](../../reference/excel/excel-add-ins-reference-overview.md)
