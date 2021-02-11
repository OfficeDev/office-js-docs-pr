---
title: Develop Office Add-ins
description: 'An introduction to developing Office Add-ins.'
ms.date: 10/14/2020
localization_priority: Priority
---

# Develop Office Add-ins

> [!TIP]
> Please review [Office Add-ins platform overview](../overview/office-add-ins.md) before reading this article.

All Office Add-ins are built upon the Office Add-ins platform. For any add-in you build, you'll need to understand important concepts like application and platform availability, Office JavaScript API programming patterns, how to specify an add-in's settings and capabilities in the manifest file, how to design the UI and user experience, and more. Core development concepts like these are covered here in the **Development lifecycle** > **Develop** section of the documentation. Review the information here before exploring the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).

## Creating an Office Add-in

You can create an Office Add-in by using the Yeoman generator for Office Add-ins or Visual Studio.

### Yeoman generator for Office Add-ins

The [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office) can be used to create a Node.js Office Add-in project that can be managed with Visual Studio Code or any other editor. The generator can create Office Add-ins for any of the following:

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word
- Excel custom functions

You can choose to create the project using HTML, CSS and JavaScript, or using Angular or React. For whichever framework you choose, you can choose between JavaScript and Typescript as well. For more information about creating add-ins with the Yeoman generator, see [Develop Office Add-ins with Visual Studio Code](../develop/develop-add-ins-vscode.md).

### Visual Studio

Visual Studio can be used to create Office Add-ins for Excel, Outlook, Word, and PowerPoint. An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript. For more information about creating add-ins with Visual Studio, see [Develop Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## Understanding the two parts of an Office Add-in

An Office Add-in consists of two parts:

- The add-in manifest (an XML file) that defines the settings and capabilities of the add-in.

- The web application that defines the UI and functionality of add-in components such as task panes, content add-ins, and dialog boxes.

The web application uses the Office JavaScript API to interact with content in the Office document where the add-in is running. Your add-in can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.

### Defining an add-in's settings and capabilities

An Office Add-in's manifest (an XML file) defines the settings and capabilities of the add-in. You'll configure the manifest to specify things such as:

- Metadata that describes the add-in (for example, ID, version, description, display name, default locale).
- Office applications where the add-in will run.
- Permissions that the add-in requires.
- How the add-in integrates with Office, including any custom UI that the add-in creates (for example, custom tabs, ribbon buttons).
- Location of images that the add-in uses for branding and command iconography.
- Dimensions of the add-in (for example, dimensions for content add-ins, requested height for Outlook add-ins).
- Rules that specify when the add-in activates in the context of a message or appointment (for Outlook add-ins only).

For detailed information about the manifest, see [Office Add-ins XML manifest](add-in-manifests.md).

### Interacting with content in an Office document

An Office Add-in can use the Office JavaScript APIs to interact with content in the Office document where the add-in is running.

#### Accessing the Office JavaScript API library

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### API models

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### API requirement sets

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

#### Exploring APIs with Script Lab

Script Lab is an add-in that enables you to explore the Office JavaScript API and run code snippets while you're working in an Office program such as Excel or Word. It's available for free via [AppSource](https://appsource.microsoft.com/product/office/WA104380862) and is a useful tool to include in your development toolkit as you prototype and verify the functionality you want in your add-in. In Script Lab, you can access a library of built-in samples to quickly try out APIs or even use a sample as the starting point for your own code.

The following one-minute video shows Script Lab in action.

[![Short video that shows Script Lab running in Excel, Word, and PowerPoint](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)

For more information about Script Lab, see [Explore Office JavaScript APIs using Script Lab](../overview/explore-with-script-lab.md).

## Extending the Office UI

An Office Add-in can extend the Office UI by using add-in commands and HTML containers such as task panes, content add-ins, or dialog boxes.

- [Add-in commands](../design/add-in-commands.md) can be used to add custom tabs, buttons, and menus to the default ribbon in Office, or to extend the default context menu that appears when users right-click text in an Office document or an object in Excel. When users select an add-in command, they initiate the task that the add-in command specifies, such as running JavaScript code, opening a task pane, or launching a dialog box.

- HTML containers like [task panes](../design/task-pane-add-ins.md), [content add-ins](../design/content-add-ins.md), and [dialog boxes](../design/dialog-boxes.md) can be used to display custom UI and expose additional functionality within an Office application. The content and functionality of each task pane, content add-in, or dialog box derives from a web page that you specify. Those web pages can use the Office JavaScript API to interact with content in the Office document where the add-in is running, and can also do other things that web pages typically do, like call external web services, facilitate user authentication, and more.

The following image shows an add-in command in the ribbon, a task pane to the right of the document, and a dialog box or content add-in over the document.

![Diagram showing add-in commands on the ribbon, a task pane, and a dialog box / content add-in in an Office document](../images/add-in-ui-elements.png)

For more information about extending the Office UI and designing the add-in's UX, see [Office UI elements for Office Add-ins](../design/interface-elements.md).

## Next steps

This article has outlined the different ways to create Office Add-ins, introduced the ways that an add-in can extend the Office UI, described the API sets, and introduced Script Lab as a valuable tool for exploring Office JavaScript APIs and prototyping add-in functionality. Now that you've explored this introductory information, consider continuing your Office Add-ins journey along the following paths.

### Create an Office Add-in

You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.yml). If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.yml).

### Learn more

Learn more about developing, testing, and publishing Office Add-ins by exploring this documentation.

> [!TIP]
> For any add-in that you build, you'll use information in the [Development lifecycle](../overview/core-concepts-office-add-ins.md) section of this documentation, along with information in the application-specific section that corresponds to the type of add-in you're building (for example, [Excel](../excel/index.yml)).

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Learn about the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
- [Design Office Add-ins](../design/add-in-design.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)