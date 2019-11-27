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

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

While you're developing your add-in, you can run the add-in on your local web server (`localhost`), but when you're ready to publish it for other users to access, you'll need to deploy the web application to a web server or web hosting service (for example, Microsoft Azure) and update the manifest to specify the URL of the deployed application. 

Complete the following steps to publish your add-in:

1. From the command line, in the root directory of your add-in project, run the following command to prepare all files for production deployment: 

    ```command&nbsp;line
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.

2. Upload the contents of the **dist** folder to the web server that'll host your add-in. You can use any type of web server or web hosting service to host your add-in.

3. In VS Code, open the add-in's manifest file, located in the root directory of the project (`manifest.xml`). Replace all occurrences of `https://localhost:3000` with the URL of the web application that you deployed to a web server in the previous step.

    > [!TIP]
    > While you can update the existing `manifest.xml` file as described here, you might instead consider preserving the existing file in its original state, and creating a copy of the file where you'll replace all instances of `https://localhost:3000` with the deployed web application's URL. Doing things this way, you'd have two versions of the manifest file: one that could be used during on-going development/testing of your add-in (referencing `localhost`) and another that could be used to publish your add-in for other users to access (referencing the deployed web application's URL).

4. Choose the method you'd like to use to [deploy and publish your Office Add-in](../publish/publish.md) your add-in, and follow the instructions to publish the manifest file.

## See also

- [5-minute quick starts](../index.md)
- [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Deploy and publish your Office Add-in](../publish/publish.md)