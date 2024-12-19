---
title: Create Office Add-in projects using Office Add-ins Development Kit for Visual Studio Code
description: Learn how to create Office Add-in projects using Office Add-ins Development Kit extension.
ms.date: 12/19/2024
ms.localizationpriority: high
---

# Create Office Add-in projects using Office Add-ins Development Kit for Visual Studio Code

The Office Add-ins Development Kit helps set up your environment, create Office Add-ins, and debug your code in a streamlined experience.

[!INCLUDE [Information about the preview status of the dev kit.](../includes/dev-kit-preview-note.md)]

[!include[Dev_kit prerequisites](../includes/dev-kit-prerequisites.md)]

## Install the development kit

You can install Office Add-ins Development Kit using **Extensions** in Visual Studio Code or install it from the Visual Studio Code Marketplace.

# [Visual Studio Code](#tab/vscode)

[!INCLUDE [Instructions to install the Office Add-ins Development Kit through VS Code.](../includes/install-dev-kit.md)]

# [Marketplace](#tab/marketplace)

1. Open the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) page in the Visual Studio Code Marketplace.
1. Select **Install** on the web page. If you're prompted that the extension requires Visual Studio Code, select **Continue**.
1. Your browser may ask you to verify that the site is trying to open Visual Studio Code. Allow Visual Studio Code to open. Visual Studio Code will then open with the Office Add-ins Development Kit extension readme displayed.
1. Select **Install** in Visual Studio Code. After successfully installing, the Office Add-ins Development Kit icon will appear in the Visual Studio Code activity bar.

---

## Create an add-in project

The Office Add-ins Development Kit offers two ways to get started with a new project: templates and samples. Templates let you choose the Office application, programming language, and starting feature. Samples are more complete projects that demonstrate a realistic scenario.

### Create an add-in from a template

Templates offer a basic starting point for your add-in. They add a minimal amount of functionality so you can get started without changing much of the existing code. The following instructions describe how to make a new project from a template using the development kit.

1. Open Visual Studio Code and select the Office Add-ins Development Kit icon in the **Activity Bar**.
1. Select **Create a New Add-in** in the extension task pane.
1. In the now-active Quick Pick menu, select the Office application for your add-in.
1. Select an add-in template from the list of available templates.
1. Select "JavaScript" or "TypeScript" as the programming language.
1. In the **Workspace folder** dialog that opens, select the folder where you want to create the project.
1. Give a name to the project (with no spaces) when prompted. Office Add-ins Development Kit will create the project with basic files and scaffolding. It then opens the project in a *second* Visual Studio Code window. You can freely close the original Visual Studio Code window.

### Create an add-in from a sample

Samples show a working add-in that solves an end-to-end scenario. Samples are most useful as learning tools to understand how features of the Office Add-ins framework work together. You can also expand a sample to fit your particular needs.

1. Open Visual Studio Code and select the Office Add-ins Development Kit icon in the **Activity Bar**.
1. Select **View Samples**.
1. Select the sample you would like to view.
1. Select the **Create** button above the now-open sample readme.
1. In the **Workspace folder** dialog that opens, select the folder where you want to create the project. The extension copies a version of the sample to that folder. It then opens the project in a *second* Visual Studio Code window. You can freely close the original Visual Studio Code window.

## Test your add-in

To understand how the add-in will work in an Office application, use the Office Add-ins Development Kit to run and debug your Office add-in in your local development environment.

> [!NOTE]
> These steps are the same as the ones listed in Visual Studio Code by the extension when you create a new project.

[!include[Dev_kit_start_debugging](../includes/dev-kit-start-debugging.md)]

[!include[Dev_kit_stop_debugging](../includes/dev-kit-stop-debugging.md)]

[!include[Dev_kit_troubleshooting](../includes/dev-kit-troubleshooting.md)]
