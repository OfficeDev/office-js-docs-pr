---
title: Develop Office Add-ins with Visual Studio
description: How to develop Office Add-ins with Visual Studio.
ms.topic: overview
ms.date: 08/18/2023
ms.localizationpriority: high
---

# Develop Office Add-ins with Visual Studio

This article describes how to use Visual Studio to develop an Office Add-in. If you've already created your add-in, you can skip ahead to the [Develop the add-in using Visual Studio](#develop-the-add-in-using-visual-studio) section.

> [!NOTE]
> Beginning with Visual Studio 2026, Office Add-in development in Visual Studio is deprecated and will be removed in a future release. Support for Office Add-in development in a different form may be added to a future version of Visual Studio.
>
> We recommend creating Office Add-in projects with the Microsoft 365 Agents Toolkit or the Yeoman Generator. For more information, see [Create Office Add-in projects using the Yeoman Generator](../develop/yeoman-generator-overview.md).

## Create the add-in project using Visual Studio

Visual Studio can be used to create Office Add-ins for Excel, Outlook, PowerPoint, and Word. An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript. To create an Office Add-in with Visual Studio, follow instructions in the quick start that corresponds to the add-in you'd like to create.

- [Excel quick start](../quickstarts/excel-quickstart-vs.md)
- [Outlook quick start](../quickstarts/outlook-quickstart-vs.md)
- [PowerPoint quick start](../quickstarts/powerpoint-quickstart-vs.md)
- [Word quick start](../quickstarts/word-quickstart-vs.md)

Visual Studio doesn't support creating Office Add-ins for OneNote or Project. To create Office Add-ins for either of these applications, you'll need to use the Yeoman generator for Office Add-ins, as described in the [OneNote quick start](../quickstarts/onenote-quickstart.md) or the [Project quick start](../quickstarts/project-quickstart.md).

## Develop the add-in using Visual Studio

Visual Studio creates a basic add-in with limited functionality. You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript, and CSS files in Visual Studio. For a high-level description of the project structure and files in the add-in project that Visual Studio creates, see the Visual Studio guidance within the quick start that you completed to create your add-in.

> [!TIP]
> Because an Office Add-in is a web application, you'll need at least basic web development skills to customize your add-in. If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

To customize your add-in, you'll need to understand concepts described in the [Core concepts > Develop](develop-overview.md) area of this documentation, as well as concepts described in the application-specific area of documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).

## Test and debug the add-in

Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform. For more information, see [Debug Office Add-ins in Visual Studio](debug-office-add-ins-in-visual-studio.md) and [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).

## Publish the add-in

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

While you're developing your add-in in Visual Studio, your add-in runs on your local web server (`localhost`). When your add-in is working as desired and you're ready to publish it for other users to access, you'll need to complete the following steps.

1. Deploy the web application to a web server or web hosting service (for example, Microsoft Azure).
2. Update the manifest to specify the URL of the deployed application.
3. Choose the method you'd like to use to [deploy your Office Add-in](../publish/publish.md), and follow the instructions to publish the manifest file.

## See also

- [Core concepts for Office Add-ins](../overview/core-concepts-office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Design Office Add-ins](../design/add-in-design.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)
