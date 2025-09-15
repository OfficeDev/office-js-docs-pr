---
title: Task pane add-ins for Project
description: Build powerful Project add-ins that streamline project management workflows, integrate with external systems, and enhance team collaboration.
ms.date: 09/15/2025
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Task pane add-ins for Project

Transform how project managers work with Microsoft Project. Task pane add-ins let you integrate project data with external systems, automate routine tasks, and create custom workflows—all without leaving the Project environment.

For an introduction to Office Add-ins, see [Office Add-ins platform overview](../overview/office-add-ins.md).

## What can you build with Project add-ins?

Project add-ins solve real-world project management challenges by connecting Project with the tools your team already uses:

### Smart scheduling assistants

- **Resource optimization**: View resource availability across multiple projects and make smarter assignments
- **Timeline coordination**: Pull milestone data from related projects to identify potential conflicts
- **Dependency management**: Automatically update schedules based on changes in connected projects

### Team collaboration tools

- **Status dashboards**: Create interactive reports that team members can update directly from Project
- **Communication hubs**: Send updates, collect feedback, and coordinate with team members without switching apps
- **Document integration**: Link project tasks to relevant documents, specifications, or design files

### Business system integration

- **ERP connectivity**: Sync budget and cost data with your accounting or ERP system
- **CRM integration**: Connect project deliverables with customer records and sales pipelines  
- **Timesheet automation**: Streamline time tracking and approval workflows
- **Reporting automation**: Generate custom reports that combine Project data with business metrics

### Industry-specific solutions

- **Construction management**: Connect with equipment scheduling, material tracking, or safety reporting systems
- **Software development**: Integrate with bug tracking, code repositories, or CI/CD pipelines
- **Marketing campaigns**: Link project milestones with campaign deadlines and deliverable tracking

## How Project add-ins work

Project add-ins are context-aware, meaning they respond to what you're currently working on. When you select a task, resource, or view in Project, your add-in can access that information and provide relevant functionality.

> [!NOTE]
> Project Professional connects seamlessly with Project Web App and Project Server, giving your add-ins access to enterprise project data. Project Standard supports add-ins but can't access Project Server integration features.

### Project editions and capabilities

**Project Professional**: Full add-in support including Project Server integration, SharePoint connectivity, and enterprise features.

**Project Standard**: Basic add-in support for local project files and standalone functionality.

## Real-world use cases

Here are some scenarios where Project add-ins add significant value:

### Scenario: Cross-project resource management

*Challenge*: A department manager needs to see resource allocation across multiple projects before making assignments.

*Solution*: An add-in that queries Project Server for all departmental projects, displays resource availability in a dashboard, and highlights potential conflicts when assigning team members.

### Scenario: Automated status reporting

*Challenge*: Weekly status reports require manually gathering data from Project, emails, and other systems.

*Solution*: An add-in that combines Project timeline data with external status updates, generates formatted reports, and emails them to stakeholders automatically.

### Scenario: Client billing integration

*Challenge*: Time tracking data in Project needs to flow into the company's billing system.

*Solution*: An add-in that captures time entries, validates them against project budgets, and exports billing data to the accounting system.

## Development options

Project add-ins use the JavaScript API, but there's currently no Project-specific API. Instead, you'll work with:

### Common Office JavaScript API

Access shared functionality like document properties, selection events, and settings storage. This API works consistently across all Office applications.

```javascript
// Example: Get the currently selected task data
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text,
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            // Process the selected task information
            processTaskData(result.value);
        }
    }
);
```

### Project Server REST APIs

For Project Professional users connected to Project Server, you can access rich project data through REST services:

- Query project information across the organization
- Access resource pools and availability
- Read and update project status
- Generate reports from Project Web App data

### Integration patterns

- **Read project data**: Access task lists, resource assignments, and timeline information
- **External data sync**: Pull information from CRM, ERP, or other business systems
- **Custom workflows**: Automate approval processes or update notifications
- **Reporting**: Generate custom dashboards and status reports

## Building your first Project add-in

Ready to get started? You can build Project add-ins using standard web development tools and techniques.

### Development setup

Create your add-in using familiar web technologies:

- **HTML** for the user interface
- **CSS** for styling and layout  
- **JavaScript** for functionality and Office API integration
- **REST APIs** for connecting to external data sources

### Add-in manifest

Your add-in needs a manifest file that configures how it appears and behaves in Project. The manifest specifies the **TaskPaneExtension** type for Project add-ins and can be shared across multiple Office applications or customized specifically for Project.

### Testing and deployment

Test your add-in thoroughly during development using the sideloading capabilities in Project. Once ready, you can distribute through:

- **File shares** for internal organizational use
- **SharePoint app catalogs** for broader internal distribution  
- **Microsoft AppSource** for public availability

For detailed guidance, see [Publish your Office Add-in](../publish/publish.md).

## Getting help and staying connected

Join the vibrant Project add-ins developer community:

- **[Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)** - Get free development resources and support
- **[Office Add-ins development overview](../develop/develop-overview.md)** - Comprehensive development guidance
- **[REST with Project Server tutorial](create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)** - Hands-on example with real Project Server integration

## Next steps

Ready to dive deeper? Start with these resources:

1. **[Build your first Project add-in](../quickstarts/project-quickstart.md)** - Get hands-on experience
2. **[Explore the sample gallery](https://developer.microsoft.com/microsoft-365/gallery/?filterBy=Project,Samples)** - See working examples
3. **[Join the community call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call)** - Connect with other developers

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Create a Project add-in that uses REST with an on-premises Project Server OData service](create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
