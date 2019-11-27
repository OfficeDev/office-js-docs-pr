---
title: Develop Office Add-ins with Visual Studio Code
description: Develop Office Add-ins with Visual Studio Code
ms.date: 11/26/2019
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

If you plan to use VS Code as your integrated development environment (IDE), you should create the Office Add-in project by using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor. 

All [5-minute quick start](../index.md) for step-by-step instructions that describe how to create an Office Add-in with the Yeoman generator.

## Develop the add-in using VS Code

When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code. 

> [!TIP]
> On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code. On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use the `code .` command to open the project folder in VS Code.

In VS Code, customize your add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files that the Yeoman generator creates by default. For a high-level description of the project structure and files in the add-in project, see the the Yeoman generator guidance within the [5-minute quick start](../index.md) that corresponds to the type of add-in you've created.

## Test and debug 

Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform. For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).

## Publish the add-in

When your add-in is working as desired and you're ready to publish it for other users to access, complete the following steps:

1. In the root directory of your add-in project, run the following command to prepare all files for production deployment: 

    ```command&nbsp;line
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.

2. Upload the contents of the **dist** folder to an [Azure Web App](https://azure.microsoft.com/en-us/services/app-service/web) or any other type of web server.

    - KB TODO: add note about not needing the manifest file?

3. KB TODO: add step -> update your manifest.xml file to point to the proper URL of where your files will be hosted --> consider creating a prod version of the manifest file.

4. KB TODO: add step -> deploy manifest file to publish add-in --> Follow one of the methods listed on [Deploy and publish your Office Add-in](/office/dev/add-ins/publish/publish) to deploy your appmanifest.xml to make your Add-in available to your users

KB TODO: add note to VS article that points back to this one.
KB TODO: update VSO work item

## See also

- [5-minute quick starts](../index.md)
- [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Deploy and publish your Office Add-in](../publish/publish.md)