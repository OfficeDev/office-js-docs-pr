---
title: Develop Office Add-ins with Visual Studio
description: How to develop Office Add-ins with Visual Studio
ms.date: 12/28/2019
localization_priority: Priority
---

# Create Office Add-ins in Visual Studio

- Pointer to quick starts for instructions about creating an add-in using VS
- Overview of content in this section; remainder of docs apply regardless of whether you've used Yo Office or Visual Studio to create your add-in (e.g., manifest docs, Office JavaScript API docs, etc.)
- Note: recommend using Yo Office instead of Visual Studio (templates are more actively maintained, supports more types of add-ins, etc.)
- Resources for getting started with web development (perhaps model after Office Scripts docs verbiage)

This article describes how to use Visual Studio to develop an Office Add-in.

> [!NOTE]
> For information about using VS Code to create an Office Add-in, see [Develop Office Add-ins with Visual Studio Code](develop-add-ins-vscode.md).

## Create the add-in project using Visual Studio

TODO

If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor. 

To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.md) that corresponds to the type of add-in you'd like to create.

[!include[Yeoman vs Visual Studio comparision](../includes/yeoman-generator-recommendation.md)]

> [!NOTE]
> Visual Studio does not support creating Office Add-ins for OneNote or Project, but you can use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create these types of add-ins.
> - To get started with an add-in for OneNote, see [Build your first OneNote add-in](../quickstarts/onenote-quickstart.md).
>
> - To get started with an add-in for Project, see [Build your first Project add-in](../quickstarts/project-quickstart.md).

## Develop the add-in using Visual Studio

TODO

When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code. 

The Yeoman generator creates a basic add-in with limited functionality. You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in Visual Studio. For a high-level description of the project structure and files in the add-in project that Visual Studio creates, see the Visual Studio guidance within the [5-minute quick start](../index.md) that corresponds to the type of add-in you've created.

## Test and debug the add-in

TODO

Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform. For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).

## Publish the add-in

TODO

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## See also

- TODO
- [5-minute quick starts](../index.md)
- [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Deploy and publish your Office Add-in](../publish/publish.md)