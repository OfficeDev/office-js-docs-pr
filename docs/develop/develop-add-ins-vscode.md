---
title: Develop Office Add-ins with Visual Studio Code
description: How to develop Office Add-ins with Visual Studio Code.
ms.topic: overview
ms.date: 01/15/2026
ms.localizationpriority: high
---

# Develop Office Add-ins with Visual Studio Code

This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.

## Prerequisites

- [Visual Studio Code](https://code.visualstudio.com/)
- A project creation tool. You have the following options.

  - The **Yeoman Generator for Office Add-ins** (also called "Yo Office"). For installation and usage instructions, see [Create Office Add-in projects using the Yeoman Generator](yeoman-generator-overview.md). With this tool, you have the option of creating add-ins that use either the add-in only manifest or the unified manifest for Microsoft 365. For more information about the differences, start with [Office add-ins manifest](add-in-manifests.md). 
  - The **Microsoft 365 Agents Toolkit for Visual Studio Code**. For installation instructions, see [Install Agents Toolkit](/microsoftteams/platform/toolkit/install-agents-toolkit). For usage instructions, see [Create Office Add-in projects with Microsoft 365 Agents Toolkit](agents-toolkit-overview.md). With this tool you can create add-ins that use the unified manifest for Microsoft 365.

  [!INCLUDE [Unified manifest support note for Office applications](../includes/unified-manifest-support-note.md)]

## Develop the add-in using VS Code

To work with the project, open the root folder of the project with VS Code.

Both tools create a basic add-in with limited functionality. You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code. For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](../index.yml) that corresponds to the type of add-in you've created.

## Test and debug the add-in

Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform and by the tool that is used to create the project. For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md) and [Create Office Add-in projects with Microsoft 365 Agents Toolkit](agents-toolkit-overview.md).

## Publish the add-in

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## See also

- [Core concepts for Office Add-ins](../overview/core-concepts-office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Design Office Add-ins](../design/add-in-design.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)
