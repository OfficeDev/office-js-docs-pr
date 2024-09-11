---
title: Create Office Add-in projects using Teams Toolkit
description: Learn how to create Office Add-in projects using Teams Toolkit.
ms.date: 09/19/2024
ms.localizationpriority: high
---

# Create Office Add-in projects with Teams Toolkit

A primary tool for developing Teams Apps is Teams Toolkit. You can create Office Add-ins with Teams Toolkit.

Add-ins created with Teams Toolkit use the [unified manifest for Microsoft 365](unified-manifest-overview.md).

   [!INCLUDE [Unified manifest support note for Office applications](../includes/unified-manifest-support-note.md)]

> [!TIP]
> There is another Visual Studio Code extension that creates Office Add-ins that use the add-in only manifest. See [Create Office Add-in projects using Office Add-ins Development Kit for Visual Studio Code](development-kit-overview.md).

   [!INCLUDE [non-unified manifest clients note](../includes/non-unified-manifest-clients.md)]

Install the latest version of Teams Toolkit into Visual Studio Code as described in [Install Teams Toolkit](/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode).

## Create an Office Add-in project

1. Open Visual Studio Code and select Teams Toolkit icon in the **Activity Bar**.

    :::image type="content" source="../images/teams-toolkit-icon.png" alt-text="Teams Toolkit icon.":::

1. Select **Create a New App**.
1. In the **New Project** drop down, select **Office Add-in**.

    :::image type="content" source="../images/teams-toolkit-create-office-add-in.png" alt-text="The five options in New Project drop down. The fifth option is called 'Office Add-in'.":::

1. In the **App Features Using an Office Add-in** drop down, select **Taskpane** or **Content Add-in**. (For simplicity, the remainder of this article assumes you selected **Taskpane**. For more information about content add-ins, see [Content Office Add-ins](../design/content-add-ins.md).)

    :::image type="content" source="../images/teams-toolkit-create-office-task-pane-capability.png" alt-text="The three options in the App Features Using an Office Add-in drop down. The first option 'Taskpane' is selected.":::

1. In the **Programming Language** drop down, select either **TypeScript** or **JavaScript**.
1. In the **Framework** drop down, select **Default** or **React**.  
1. In the **Workspace Folder** dialog that opens, select the folder where you want to create the project.
1. Give a name to the project (with no spaces) when prompted. Teams Toolkit will create the project with basic files and scaffolding. It will then open the project *in a second Visual Studio Code window*. Close the original Visual Studio Code window.

   > [!NOTE]
   > The project that is generated is configured to be installable on Excel, Outlook, PowerPoint, and Word. You can edit the manifest and source files as needed to change which Office applications are supported.

1. In the Visual Studio Code **TERMINAL** navigate to the root of the project and run `npm install`.
1. After the installation completes, verify that you can sideload your add-in from Visual Studio Code. The steps to sideload vary depending on the Office application on which you want to test the add-in.

### Sideload in Excel, PowerPoint, or Word

1. Select **View** | **Run** in Visual Studio Code. In the **RUN AND DEBUG** drop down menu, select one of these options:
 
    - **Excel Desktop (Edge Chromium)**
    - **PowerPoint Desktop (Edge Chromium)**
    - **Word Desktop (Edge Chromium)**

1. Press F5. The project builds and a Node dev-server window opens. This process may take a couple of minutes. Eventually, the desktop version of the Office application you selected opens.

    > [!NOTE]
    > If this is the first time that you have sideloaded an Office add-in on your computer (or the first time in over a month), you may be prompted to delete an old certificate and/or to install a new one. Agree to both prompts.

1. A **Contoso Add-in** tab with two buttons will appear on the **Home** ribbon. Use one button to perform an action in the open Office document. Use the other to open the add-in's task pane.

    > [!NOTE]
    > Regardless of which button you select, a **WebView Stop On Load** prompt appears. Select **OK**.

1. To stop debugging and uninstall the add-in, select **Run** | **Stop Debugging** in Visual Studio Code.

### Sideload in Outlook

1. Ensure that your account in your Microsoft 365 developer tenancy is also an email account in desktop Outlook. If it isn't, follow the guidance in [Add an email account to Outlook](https://support.microsoft.com/office/add-an-email-account-to-outlook-e9da47c4-9b89-4b49-b945-a204aeea6726).
1. **Close Outlook desktop**.
1. In Visual Studio Code, open Teams Toolkit.
1. In the **ACCOUNTS** section, verify that you're signed into Microsoft 365.
1. Select **View** | **Run** in Visual Studio Code. In the **RUN AND DEBUG** drop down menu, select the option, **Outlook Desktop (Edge Chromium)**, and then press F5. The project builds and a Node dev-server window opens. This process may take a couple of minutes and then Outlook desktop will open.

    > [!NOTE]
    > If this is the first time that you have sideloaded an Office add-in on your computer (or the first time in over a month), you may be prompted to delete an old certificate and/or to install a new one. Agree to both prompts.

1. Open the **Inbox** *of your Microsoft 365 account identity* and open any message. A **Contoso Add-in** tab with two buttons will appear on the **Home** ribbon (or the **Message** ribbon, if you have opened the message in its own window).
1. Click the **Show Taskpane** button and a task pane opens. Click the **Perform an action** button and a small notification appears near the top of the message.

    > [!NOTE]
    > Regardless of which button you select, a **WebView Stop On Load** prompt appears. Select **OK**.

1. To stop debugging and uninstall the add-in, select **Run** | **Stop Debugging** in Visual Studio Code.

## Developing your project

Now you can change and develop the project. In places where the guidance in the Office Add-ins documentation branches depending on what type of manifest is being used, be sure to follow the guidance for the unified manifest.
