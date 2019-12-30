---
title: Develop Office Add-ins with Visual Studio
description: How to develop Office Add-ins with Visual Studio
ms.date: 12/31/2019
localization_priority: Priority
---

# Develop Office Add-ins with Visual Studio

This article describes how to use Visual Studio to develop an Office Add-in. If you've already created your add-in, you can skip ahead to the [Develop the add-in using Visual Studio](#develop-the-add-in-using-visual-studio) section.

> [!NOTE]
> As an alternative to using Visual Studio, you may choose to use the Yeoman generator for Office Add-ins and VS Code to create an Office Add-in. For more information about this choice, see [Creating an Office Add-in](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in).

## Create the add-in project using Visual Studio

Visual Studio can be used to create Office Add-ins for Excel, Outlook, Word, and PowerPoint. An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript. To create an Office Add-in with Visual Studio, follow instructions in the quick start that corresponds to the add-in you'd like to create:

- [Excel quick start](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Outlook quick start](/outlook/add-ins/quick-start?context=office/dev/add-ins/context&tabs=visualstudio)
- [Word quick start](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [PowerPoint quick start](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

Visual Studio doesn't support creating Office Add-ins for OneNote or Project. To create Office Add-ins for either of these hosts, you'll need to use the Yeoman generator for Office Add-ins, as described in the [OneNote quick start](../quickstarts/onenote-quickstart.md) or the [Project quick start](../quickstarts/project-quickstart.md).

## Develop the add-in using Visual Studio

Visual Studio creates a basic add-in with limited functionality. You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript, and CSS files in Visual Studio. For a high-level description of the project structure and files in the add-in project that Visual Studio creates, see the Visual Studio guidance within the quick start that you completed to create your add-in. 

> [!TIP]
> Because an Office Add-in is a web application, you'll need at least basic web development skills to customize your add-in. If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

To customize your add-in, you'll need to understand concepts described in the [Core concepts > Develop](develop-overview.md) area of this documentation, as well as concepts described in the host-specific area of documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.md)). 

## Test and debug the add-in

Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform. For more information, see [Debug Office Add-ins in Visual Studio](debug-office-add-ins-in-visual-studio.md) and [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).

## Publish the add-in

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

While you're developing your add-in in Visual Studio, your add-in runs on your local web server (`localhost`). When your add-in is working as desired and you're ready to publish it for other users to access, you'll need to complete the following steps:

1. Deploy the web application to a web server or web hosting service (for example, Microsoft Azure).
2. Update the manifest to specify the URL of the deployed application. 
3. Choose the method you'd like to use to [deploy your Office Add-in](../publish/publish.md), and follow the instructions to publish the manifest file.

## See also

- [Building Office Add-ins](../overview/office-add-ins-fundamentals.md)
- [Core concepts for Office Add-ins](../overview/core-concepts-office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Design Office Add-ins](../design/add-in-design.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)