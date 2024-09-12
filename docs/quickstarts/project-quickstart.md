---
title: Build your first Project task pane add-in
description: Learn how to build a simple Project task pane add-in by using the Office JS API.
ms.date: 12/11/2023
ms.service: project
ms.localizationpriority: high
---

# Build your first Project task pane add-in

In this article, you'll walk through the process of building a Project task pane add-in.

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 or later on Windows

## Create the add-in

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project`
- **Choose a script type:** `JavaScript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `Project`

![The prompts and answers for the Yeoman generator in a command line interface.](../images/yo-office-project.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

## Explore the project

The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.

- The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application. In this quick start, the code sets the `Name` field and `Notes` field of the selected task of a project.

## Try it out

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Start the local web server.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    Run the following command in the root directory of your project. When you run this command, the local web server will start.

    ```command&nbsp;line
    npm run dev-server
    ```

1. In Project, create a simple project plan.

1. Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

1. Select a single task within the project.

1. At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.

    ![The Project application with the task pane add-in loaded.](../images/project-quickstart-addin-1.png)

1. When you want to stop the local web server and uninstall the add-in, follow these instructions:

    - To stop the server, run the following command.

        ```command&nbsp;line
        npm stop
        ```

    - To uninstall the sideloaded add-in, see [Remove a sideloaded add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md#remove-a-sideloaded-add-in).

## Next steps

Congratulations, you've successfully created a Project task pane add-in! Next, learn more about the capabilities of a Project add-in and explore common scenarios.

> [!div class="nextstepaction"]
> [Project add-ins](../project/project-add-ins.md)

[!include[The common troubleshooting section for all Yo Office quick starts](../includes/quickstart-troubleshooting-yo.md)]

## See also

- [Develop Office Add-ins](../develop/develop-overview.md)
- [Core concepts for Office Add-ins](../overview/core-concepts-office-add-ins.md)
- [Using Visual Studio Code to publish](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
