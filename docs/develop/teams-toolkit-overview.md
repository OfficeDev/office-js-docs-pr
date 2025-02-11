---
title: Create Office Add-in projects using Teams Toolkit
description: Learn how to create Office Add-in projects using Teams Toolkit.
ms.date: 04/12/2024
ms.localizationpriority: high
---

# Create Office Add-in projects with Teams Toolkit

A primary tool for developing Teams Apps is Teams Toolkit. You can create Office Add-ins with Teams Toolkit, with the following restrictions.

- Add-ins created with Teams Toolkit use the [unified manifest for Microsoft 365](unified-manifest-overview.md).
- Only Outlook add-ins can be created at this time. We're working hard to enable support in Teams Toolkit for add-ins to other Office applications and platforms.

   [!INCLUDE [non-unified manifest clients note](../includes/non-unified-manifest-clients.md)]

Install the latest version of Teams Toolkit into Visual Studio Code as described in [Install Teams Toolkit](/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode).

## Create an Outlook Add-in project

1. Open Visual Studio Code and select Teams Toolkit icon in the **Activity Bar**.

    :::image type="content" source="../images/teams-toolkit-icon.png" alt-text="Teams Toolkit icon.":::

1. Select **Create a new app**.
1. In the **New Project** drop down, select **Outlook add-in**.

    :::image type="content" source="../images/teams-toolkit-create-outlook-add-in.png" alt-text="The four options in New Project drop down. The fourth option is called 'Outlook add-in'.":::

1. In the **App Features Using an Outlook Add-in** drop down, select **Taskpane**.

    :::image type="content" source="../images/teams-toolkit-create-outlook-task-pane-capability.png" alt-text="The two options in the App Features Using an Outlook Add-in drop down. The first option 'Taskpane' is selected.":::

1. In the **Workspace folder** dialog that opens, select the folder where you want to create the project.
1. Give a name to the project (with no spaces) when prompted. Teams Toolkit will create the project with basic files and scaffolding. It will then open the project *in a second Visual Studio Code window*. Close the original Visual Studio Code window.
1. In the Visual Studio Code **TERMINAL** navigate to the root of the project and run `npm install`.
1. Before you make changes to the project, verify that you can sideload your Outlook add-in from Visual Studio Code. Use the following steps:
    1. Ensure that your account in your Microsoft 365 developer tenancy is also an email account in desktop Outlook. If it isn't, follow the guidance in [Add an email account to Outlook](https://support.microsoft.com/office/add-an-email-account-to-outlook-e9da47c4-9b89-4b49-b945-a204aeea6726).
    1. **Close Outlook desktop**.
    1. In Visual Studio Code, open Teams Toolkit.
    1. In the **ACCOUNTS** section, verify that you're signed into Microsoft 365.
    1. Select **View** | **Run** in Visual Studio Code. In the **RUN AND DEBUG** drop down menu, select the option, **Outlook Desktop (Edge Chromium)**, and then press <kbd>F5</kbd>. The project builds and a Node dev-server window opens. This process may take a couple of minutes. Eventually, Outlook desktop will open.
    1. Open the **Inbox** *of your Microsoft 365 account identity* and open any message. A **Contoso Add-in** tab with two buttons will appear on the **Home** ribbon (or the **Message** ribbon, if you have opened the message in its own window).
    1. Click the **Show Taskpane** button and a task pane opens. Click the **Perform an action** button and a small notification appears near the top of the message.
    1. To stop debugging and uninstall the add-in, select **Run** | **Stop Debugging** in Visual Studio Code.

Now you can change and develop the project. In places where the guidance in the Office Add-ins documentation branches depending on what type of manifest is being used, be sure to follow the guidance for the unified manifest.
