---
title: Develop Office Add-ins with Visual Studio Code
description: How to develop Office Add-ins with Visual Studio Code.
ms.topic: overview
ms.date: 08/15/2024
ms.localizationpriority: high
---

# Develop Office Add-ins with Visual Studio Code

This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.

> [!NOTE]
> For information about using Visual Studio to create an Office Add-in, see [Develop Office Add-ins with Visual Studio](develop-add-ins-visual-studio.md).

## Prerequisites

- [Visual Studio Code](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Create the add-in project using the Yeoman generator

If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](yeoman-generator-overview.md). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor.

To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.yml) that corresponds to the type of add-in you'd like to create.

## Develop the add-in using VS Code

When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code.

[!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

The Yeoman generator creates a basic add-in with limited functionality. You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code. For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](../index.yml) that corresponds to the type of add-in you've created.

### Create the add-in project using the Office Add-ins Development Kit (preview)

The [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) is a Visual Studio Code extension that allows you to create new projects directly from VS Code. For information on installing the extension and creating projects from templates and samples, see [Create Office Add-in projects using Office Add-ins Development Kit for Visual Studio Code](development-kit-overview.md).

[!INCLUDE [Information about the preview status of the dev kit.](../includes/dev-kit-preview-note.md)]

## Test and debug the add-in

Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform. For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).

## Publish the add-in

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## See also

- [Core concepts for Office Add-ins](../overview/core-concepts-office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Design Office Add-ins](../design/add-in-design.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)
