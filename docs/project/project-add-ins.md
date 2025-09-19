---
title: Task pane add-ins for Project
description: Learn about task pane add-ins for Project.
ms.date: 09/16/2025
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Task pane add-ins for Project

Build custom task pane add-ins to extend Project with web integrations and streamlined workflows. Project add-ins help project managers consolidate critical information, manage resources, and collaborate with teams, all without leaving Project.

For an introduction to Office Add-ins, see [Office Add-ins platform overview](../overview/office-add-ins.md).

> [!NOTE]
> Project Professional supports task pane add-ins that can access Project on the web, on-premises Project Server installations, and SharePoint (both on-premises and online). Project Standard doesn't support direct integration with Project Server data or SharePoint task lists synchronized with Project Server.

## Project add-in scenarios

Project task pane add-ins are context-sensitive, meaning they can respond to your currently selected task, resource, view, or Gantt chart data. This creates opportunities for targeted, relevant functionality that enhances your project management workflow.

Here are the primary ways project managers use Project add-ins:

### External system integration

**Enterprise data connections**: Connect Project with external systems like ERP applications, timesheet systems, or accounting software. Pull relevant data into Project or push project updates to other business systems.

**Custom workflows**: Build approval processes and status update workflows that span multiple systems while keeping Project as the central planning tool.

### Project Server and SharePoint integration

**Project Server data access**: Use Project Server's REST APIs and OData services to access portfolio data, resource information, and reporting data. Integrate selected task or resource information with broader organizational project data.

**SharePoint connectivity**: Connect with SharePoint libraries, lists, and workflows. Access project templates, document libraries, and team collaboration spaces directly from Project.

### Data visualization and reporting

**Custom dashboards**: Create specialized views of project data by combining Project information with external data sources. Build reports and visualizations tailored to your organization's needs.

**Real-time data display**: Show live data from other systems alongside your project plan, helping you make informed decisions without switching applications.

> [!TIP]
> For a detailed example of Project Server integration, see [Create a Project add-in that uses REST with an on-premises Project Server OData service](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).

## Build your Project add-in

Project add-ins use the Office JavaScript API to interact with Project data and integrate with external services. While there's no Project-specific JavaScript API, you can use the [Common API](/javascript/api/office) to create add-ins.

### Development approach

You have flexibility in how you build your Project add-in:

- **Simple approach**: Create an HTML webpage with JavaScript, CSS, and REST queries using any text editor.
- **Framework-based**: Use modern web frameworks like React, Angular, or Vue.js for more complex user interfaces.
- **Server-side**: Build with ASP.NET, Node.js, PHP, or other server technologies for backend integration.

### Required components

Every Project add-in needs two key components:

1. **Web application**: Your HTML, CSS, and JavaScript files that provide the user interface and functionality.
2. **Manifest file**: An XML configuration file that tells Project how to integrate your add-in.

The manifest file specifies the `TaskPaneExtension` type and can be shared across multiple Office applications or created specifically for Project. Learn more about manifests in the [Office Add-ins platform overview](../overview/office-add-ins.md).

### Development best practices

- **Test continuously**: Sideload and test your add-in frequently during development to catch issues early
- **Start simple**: Begin with basic functionality and gradually add complexity
- **Use familiar web technologies**: Leverage your existing HTML, CSS, and JavaScript skills

> [!TIP]
> Learn about testing and sideloading techniques in [Test Office Add-ins](../testing/test-debug-office-add-ins.md).

## Share your Project add-in

Once you've built your Project add-in, you have several options for distribution:

- **File share**: Share manifest files through network file shares for small team or departmental deployments
- **SharePoint app catalog**: Deploy through your organization's SharePoint app catalog for enterprise distribution
- **AppSource**: Publish to Microsoft AppSource to reach Project users worldwide

Each distribution method has different benefits depending on your target audience and organizational requirements. Learn more about your options in [Publish your Office Add-in](../publish/publish.md).

## Get started

Ready to being building your first Project add-in? Build an add-in in minutes by following with this quick start.

> [!div class="nextstepaction"]
> [Get started with your first Project add-in](../quickstarts/project-quickstart.md)

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
- [Developing Office Add-ins](../develop/develop-overview.md)
- [Create a Project add-in that uses REST with an on-premises Project Server OData service](create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
