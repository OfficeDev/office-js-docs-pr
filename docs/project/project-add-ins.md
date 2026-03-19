---
title: Task pane add-ins for Project
description: Learn about task pane add-ins for Project.
ms.date: 03/19/2026
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Task pane add-ins for Project

Build custom task pane add-ins to extend Project with web integrations and streamlined workflows. Project add-ins help project managers consolidate critical information, manage resources, and collaborate with teams, all without leaving Project.

For an introduction to Office Add-ins, see [Office Add-ins platform overview](../overview/office-add-ins.md).

> [!NOTE]
> Project Professional supports task pane add-ins. Project Standard also supports task pane add-ins but with more limited capabilities.
>
> [!IMPORTANT]
> Project Server Subscription Edition (the current supported on-premises version) has removed several integration features available in earlier versions, including the ProjectData OData service. Project Online is retiring on September 30, 2026. For current Project add-in capabilities, see the [Common API documentation](/javascript/api/office).

## Project add-in scenarios

Project task pane add-ins are context-sensitive, meaning they can respond to your currently selected task, resource, view, or Gantt chart data. This creates opportunities for targeted, relevant functionality that enhances your project management workflow.

Here are the primary ways project managers use Project add-ins:

### External system integration

**Custom workflows**: Build approval processes and status update workflows that integrate with external systems while keeping Project as the central planning tool.

> [!NOTE]
> Project add-ins have limited built-in data query capabilities. The Common API for Project primarily supports task and resource selection events. Advanced data integration scenarios that were previously enabled by the ProjectData OData service (removed in Project Server Subscription Edition) are no longer available through the Office JavaScript API.

## Build your Project add-in

Project add-ins use the Office JavaScript API to interact with Project data and integrate with external services. While there's no Project-specific JavaScript API, you can use the [Common API](/javascript/api/office) to create add-ins.

> [!NOTE]
> The Common API support for Project is limited compared to other Office applications. It primarily provides access to task, resource, and view selection events, along with basic project field data. Complex data queries, reporting, and aggregation scenarios require alternative approaches outside of the Office JavaScript API.

### Development approach

You have flexibility in how you build your Project add-in:

- **Simple approach**: Create an HTML webpage with JavaScript, and CSS using any text editor.
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
- **Microsoft Marketplace**: Publish to Microsoft Marketplace to reach Project users worldwide

Each distribution method has different benefits depending on your target audience and organizational requirements. Learn more about your options in [Publish your Office Add-in](../publish/publish.md).

## Get started

Ready to begin building your first Project add-in? Build an add-in in minutes by following this quick start.

> [!div class="nextstepaction"]
> [Get started with your first Project add-in](../quickstarts/project-quickstart.md)

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
- [Developing Office Add-ins](../develop/develop-overview.md)
