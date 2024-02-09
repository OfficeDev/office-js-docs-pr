---
title: Task pane add-ins for Project
description: Learn about task pane add-ins for Project.
ms.date: 01/23/2024
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Task pane add-ins for Project

Project Standard 2013 and Project Professional 2013 (version 15.1 or higher) both include support for task pane add-ins. You can run general task pane add-ins that are developed for Word or Excel. You can also develop custom add-ins that handle selection events in Project and integrate task, resource, view, and other cell-level data in a project with SharePoint lists, SharePoint Add-ins, Web Parts, web services, and enterprise applications.

For an introduction to Office Add-ins, see [Office Add-ins platform overview](../overview/office-add-ins.md).

## Add-in scenarios for Project

Project managers can use Project task pane add-ins to help with project management activities. Instead of leaving Project and opening another application to search for frequently used information, project managers can directly access the information within Project. The content in a task pane add-in can be context-sensitive, based on the selected task, resource, view, or other data in a cell in a Gantt chart, task usage view, or resource usage view.

> [!NOTE]
> With Project Professional 2013, you can develop task pane add-ins that access Project on the web, on-premises installations of Project Server 2013, and on-premises or online SharePoint 2013. Project Standard 2013 does not support direct integration with Project Server data or SharePoint task lists that are synchronized with Project Server.

Add-in scenarios for Project include the following:

- **Project scheduling** View data from related projects that can affect scheduling. A task pane add-in can integrate relevant data from other projects in Project Server 2013. For example, you can view the departmental collection of projects and milestone dates, or view specified data from other projects that are based on a selected custom field.

- **Resource management** View the complete resource pool in Project Server 2013 or a subset based on specified skills, including cost data and resource availability, to help select appropriate resources.

- **Statusing and approvals** Use a web application in a task pane add-in to update or view data from an external enterprise resource planning (ERP) application, timesheet system, or accounting application. Or, create a custom status approval Web Part that can be used within both Project Web App and Project Professional 2013.

- **Team communication** Communicate with team members and resources directly from a task pane add-in, within the context of a project. Or, easily maintain a set of context-sensitive notes for yourself as you work in a project.

- **Work packages** Search for specified kinds of project templates within SharePoint libraries and online template collections. For example, find templates for construction projects and add them to your Project template collection.

- **Related items** View metadata, documents, and messages that are related to specific tasks in a project plan. For example, you can use Project Professional 2013 to manage a project that was imported from a SharePoint task list, and still synchronize the task list with changes in the project. A task pane add-in can show additional fields or metadata that Project did not import for tasks in the SharePoint list.

- **Use the Project Server object models** Use the GUID of a selected task with methods in the Project Server Interface (PSI) or the client-side object model (CSOM) of Project Server. For example, the web application for an add-in can read and update the statusing data of a selected task and resource, or integrate with an external timesheet application.

- **Get reporting data** Use Representational State Transfer (REST), JavaScript, or LINQ queries to find related information for a selected task or resource in the OData service for reporting tables in Project Web App. Queries that use the OData service can be done with an online or an on-premises installation of Project Server 2013.

    For example, see [Create a Project add-in that uses REST with an on-premises Project Server OData service](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).

## Develop Project add-ins

Project supports add-ins made with the JavaScript API, but there's currently no JavaScript API designed specifically for interacting with Project. You can use the [Common API](/javascript/api/office) to create Project add-ins. For information about the Common API, see [Office JavaScript API object model](../develop/office-javascript-api-object-model.md).

To create an add-in, you can use a simple text editor to create an HTML webpage and related JavaScript files, CSS files, and REST queries. In addition to an HTML page or a web application, an add-in requires an XML manifest file for configuration. Project can use a manifest file that includes a **type** attribute that is specified as **TaskPaneExtension**. The manifest file can be used by multiple Office 2013 client applications, or you can create a manifest file that is specific for Project 2013. For more information, see the  _Development basics_ section in [Office Add-ins platform overview](../overview/office-add-ins.md).

Be sure to test your add-in as you develop it. Learn about testing and sideloading add-in in the article [Test Office Add-ins](../testing/test-debug-office-add-ins.md).

## Distribute Project add-ins

You can distribute add-ins through a file share, an app catalog in a SharePoint library, or AppSource. For more information, see [Publish your Office Add-in](../publish/publish.md).

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Learn about the Microsoft 365 Developer Program](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-)
- [Developing Office Add-ins](../develop/develop-overview.md)
- [Create a Project add-in that uses REST with an on-premises Project Server OData service](create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
