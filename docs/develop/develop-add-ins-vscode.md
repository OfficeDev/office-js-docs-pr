---
title: Develop Office Add-ins with Visual Studio Code
description: Develop Office Add-ins with Visual Studio Code
ms.date: 11/27/2019
localization_priority: Priority
---

# Develop Office Add-ins with Visual Studio Code

This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in. 

> [!NOTE]
> For information about using Visual Studio to create an Office Add-in, see [Create and debug Office Add-ins in Visual Studio](create-and-debug-office-add-ins-in-visual-studio.md).

## Prerequisites

- [Visual Studio Code](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Create the add-in project using the Yeoman generator

If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor. 

To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.md) that corresponds to the type of add-in you'd like to create.

## Develop the add-in using VS Code

When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code. 

> [!TIP]
> On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code. On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use that command to open the project folder in VS Code.

The Yeoman generator creates a basic add-in with limited functionality. You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code. For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the the Yeoman generator guidance within the [5-minute quick start](../index.md) that corresponds to the type of add-in you've created.

## Test and debug 

Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform. For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).

## Publish the add-in

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## See also

- [5-minute quick starts](../index.md)
- [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Deploy and publish your Office Add-in](../publish/publish.md)