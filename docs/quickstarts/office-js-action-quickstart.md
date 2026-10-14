---
title: Build your first Office JavaScript Library action
description: Learn how to build a simple Office.js action that Copilot in Excel can invoke.
ms.date: 10/19/2026
ms.topic: how-to
ms.service: microsoft-365
ms.localizationpriority: high
---

# Build your first Office JavaScript Library action

In this article, you'll walk through the process of building a simple Office JavaScript Library (Office.js) action that Copilot can invoke to perform actions on the content of an Office document. The project also includes a task pane.

## Prerequisites

- Node.js (the latest Active LTS version). Visit the [Node.js](https://nodejs.org/en/download) site to download and install the right version for your operating system.
- **Microsoft 365 Agents Toolkit**. For installation instructions, see [Install Microsoft 365 Agents Toolkit](/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode).

## Create the project

1. In Visual Studio Code, open Microsoft 365 Agents Toolkit.

    :::image type="content" source="../images/agent-toolkit-start-panel.png" alt-text="Agents Toolkit start panel.":::

1. Select **Create a New Agent/App**.
1. In the **New Project** list, select **Office Add-in**.

    :::image type="content" source="../images/atk-toolkit-new-project-list.png" alt-text="Agents Toolkit new project list.":::

1. On the **Select a capability** list, select **Create Office JavaScript Library Actions**.

    :::image type="content" source="../images/agent-toolkit-add-in-capability-list.png" alt-text="A dropdown list of add-in capabilities including, Task pane, Custom Function and Shortcut, Create Office JavaScript Library Actions, and Upgrade an Existing Office Add-in.":::

1. In the refined capability list that opens, select **New Office JavaScript Library Actions**.

    :::image type="content" source="../images/agent-toolkit-refined-add-in-capability-list.png" alt-text="A dropdown list of two choices: New Office JavaScript Library Actions, and Extend an Existing Office Add-in.":::

1. In the **Workspace Folder** control that opens, choose a folder for the project.
1. In the **Application Name** text box, enter "My Office.js Action".
1. The project opens in a new Visual Studio Code window. Close the original Visual Studio Code window.
1. Open the Agents Toolkit, and in the **ACCOUNTS** section, ensure that you're logged into your Microsoft 365 tenancy and that both **Custom App Upload** and **Copilot Access** are enabled.

## Run the project

1. Close all Office applications.
1. Select **View | Run** in Visual Studio Code. In the **RUN AND DEBUG** dropdown menu, select **Excel Desktop (Edge Chromium)**.

1. Press <kbd>F5</kbd>. The project builds and a Node dev-server window opens. This process may take a couple of minutes. Eventually, the Office application opens.

> [!NOTE]
> If this is the first time that you've sideloaded a project on your computer (or the first time in over a month), you may be prompted to delete an old certificate and/or to install a new one. Agree to both prompts.

1. Verify that the the **Show Task Pane** button appears on the **Home** ribbon. If the button doesn't appear, select the **Add-ins** button, and then select **My Office.js Action** on the flyout that opens.
1. Create a table with the following data. 

| Principal | Annual Interest | Expected Sale |
|----|----|----|
| 100 | 0.05 | 2027 |
| 300 | 0.1 | 2028 |
| 250 | 0.03 | 2027 |

You can start working with either the task pane or Copilot. (When both have been started, then they each have an icon tab on the right side of the task pane so you can switch between them.)

### Start the task pane

1. Select the **Show Task Pane** button on the ribbon. 
1. Test the project by pressing the **Run** link in the task pane, or selecting the **Perform an action** button.

### Start Copilot

1. Select the dynamic action button for Copilot Chat on the surface of the open document.

    :::image type="content" source="../images/copilot-dynamic-action-button.png" alt-text="The dynamic action button for Copilot: a round button with the Copilot symbol. There is no text on the button.":::

1. Enter the following prompt, or something synonymous with it, but replace the placeholder {{address}} with the range of the table, for example, `A1:C4`: `Apply Contoso's required formatting to the range {{address}}'.

## Shut down the session completely

It's important to shut down the debugging session and uninstall the project completely to avoid subtle problems when artifacts are cached. Use the following steps.

1. In Visual Studio Code, open the **Run** menu and select **Stop debugging**, or press <kbd>Shift</kbd>+<kbd>F5</kbd>. Due to a bug we're working on, this action doesn't always completely shut down the server, close the Office application, and uninstall the project. So, carry out the remaining steps.
1. Close the Office application if it's still running.
1. If the dev server is still running, shutting down the server depends on what window it's running in.

   - If the web server is running in a separate window from Visual Studio Code, open a command prompt or the Visual Studio Code **TERMINAL**. In the root of the project, run `npm run stop`.
   - If the web server is running in the Visual Studio Code **TERMINAL**, give the window focus and press <kbd>Ctrl</kbd>+<kbd>C</kbd>. Choose "Y" in response to the prompt to end the process.

1. Clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select the gear icon in the upper right.
1. If the **My Office.js Action** project is in the list of apps, select the arrow head to the left of the name to expand its row.
1. Select the trash can icon near the right end of the row, and then select **Remove** in the prompt.

## Troubleshooting

See [Troubleshooting Office.js actions in Copilot](../design/office-js-action-overview.md#troubleshooting-officejs-actions-in-copilot).

## Next steps

- [Add an Office.js action to Copilot](../develop/office-js-action.md).

## See also

- [Combine Copilot with Office.js actions](../design/office-js-action-overview.md)
- [Action file schema 1.0 for Office.js actions](action-file-schema-1.0.md)
