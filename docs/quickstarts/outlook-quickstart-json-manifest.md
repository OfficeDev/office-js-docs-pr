---
title: Build an Outlook add-in with a Teams manifest (preview)
description: Learn how to build a simple Outlook task pane add-in with a JSON manifest.
ms.date: 05/24/2022
ms.prod: outlook
ms.localizationpriority: high
---

# Build an Outlook add-in with a Teams manifest (preview)

In this article, you'll walk through the process of building an Outlook task pane add-in that displays at least one property of a selected message. This add-in will use a preview version of the JSON-formatted manifest that Teams extensions, like custom tabs and messaging extensions, use.  For more information about this manifest, see [Teams manifest for Office Add-ins (preview)](../develop/json-manifest-overview.md).

The new manifest is available for preview and we encourage experienced add-in developers to experiment with it. It should not be used in production add-ins. The preview is only supported on subscription Office on Windows. 

> [!TIP]
> If you want to build an Outlook add-in using the XML manifest, see [Build your first Outlook add-in](outlook-quickstart.md).

## Create the add-in

You can create an Office Add-in with a JSON manifest by using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md). The Yeoman generator creates a Node.js project that can be managed with Visual Studio Code or any other editor.

### Prerequisites

[.NET runtime](https://dotnet.microsoft.com/download/dotnet/6.0/runtime) for Windows. One of the tools used in the preview runs on .NET.

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

[!INCLUDE [Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- [Visual Studio Code (VS Code)](https://code.visualstudio.com/) or your preferred code editor

- Outlook on Windows (connected to a Microsoft 365 account)

### Create the add-in project

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Choose a project type** - `Outlook Add-in with Teams Manifest (Developer preview)`

    - **What do you want to name your add-in?** - `Outlook Add-in with Teams Manifest`

    SCREENSHOT HERE

    After you complete the wizard, the generator will create the project and install supporting Node components.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Navigate to the root folder of the web application project.

    ```command&nbsp;line
    cd "Outlook Add-in with Teams Manifest"
    ```

### Explore the project

The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in which also has a custom ribbon button that [executes a command](../develop/create-addin-commands?branch=outlook-json-manifest#step-5-add-the-functionfile-element).

- The **./manifest.json** file in the root directory of the project defines the settings and capabilities of the add-in.
- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Outlook.
- The **./src/command/command.html** file contains the HTML `<script>` tag that loads the JavaScript file that is transpiled from the command.ts file.
- The **./src/command/command.ts** file contains the Office JavaScript API code that executes when the custom ribbon button is selected.

### Update the code

1. Open your project in VS Code or your preferred code editor.
   [!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

1. Open the file **./src/taskpane/taskpane.html** and replace the entire **\<main\>** element (within the **\<body\>** element) with the following markup. This new markup adds a label where the script in **./src/taskpane/taskpane.js** will write data.

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function. This code uses the Office JavaScript API to get a reference to the current message and write its **subject** property value to the task pane.

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### Try it out

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. Run the following command in the root directory of your project. When you run this command, the local web server starts and your add-in will be [sideloaded](../outlook/sideload-outlook-add-ins-for-testing.md).

    ```command&nbsp;line
    npm start
    ```

1. In Outlook, view a message in the [Reading Pane](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0), or open the message in its own window.

1. Choose the **Home** tab (or the **Message** tab if you opened the message in a new window), and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![Screenshot showing a message window in Outlook with the add-in ribbon button highlighted.](../images/quick-start-button-1.png)

    > [!NOTE]
    > If you receive the error "We can't open this add-in from localhost" in the task pane, follow the steps outlined in the [troubleshooting article](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).

1. When prompted with the **WebView Stop On Load** dialog box, select **OK**.

    [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

1. Scroll to the bottom of the task pane and choose the **Run** link to write the message subject to the task pane.

    ![Screenshot showing the add-in's task pane with the Run link highlighted.](../images/quick-start-task-pane-2.png)

    ![Screenshot of the add-in's task pane displaying message subject.](../images/quick-start-task-pane-3.png)

### Add a custom button to the ribbon





