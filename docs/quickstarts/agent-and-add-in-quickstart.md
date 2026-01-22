---
title: Build your first add-in as a Copilot skill
description: Learn how to build a simple Copilot agent that has an Excel add-in as a skill.
ms.date: 01/25/2026
ms.topic: how-to
ms.service: microsoft-365
ms.localizationpriority: high
---

# Build your first add-in as a Copilot skill

In this article, you'll walk through the process of building a simple Copilot agent that can perform actions on the content of an Office document. The app also includes a task pane add-in.

## Knowledge prerequisites

- A basic understanding of declarative agents in Microsoft 365 Copilot. If you aren't familiar with them already, we recommend reading [Declarative agents for Microsoft 365 Copilot overview](/microsoft-365-copilot/extensibility/overview-declarative-agent).

## Software prerequisites

- All the prerequisites listed at [Create declarative agents using Microsoft 365 Agent Toolkit](/microsoft-365-copilot/extensibility/build-declarative-agents).
- **Microsoft 365 Agents Toolkit**. For installation instructions, see [Install Microsoft 365 Agents Toolkit](/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode).

## Create the project

1. In Visual Studio Code, open Microsoft 365 Agents Toolkit. 

    :::image type="content" source="../images/agent-toolkit-start-panel.png" alt-text="Agents Toolkit start panel.":::

1. Select **Create a New Agent/App**.
1. In the **New Project** list, select **Office Add-in**.

    :::image type="content" source="../images/atk-toolkit-new-project-list.png" alt-text="Agents Toolkit new project list.":::

1. On the **Select a capability** list, select **Create Declarative Agent with Office Add-in Action**.

    :::image type="content" source="../images/agent-toolkit-add-in-capability-list.png" alt-text="A dropdown list of add-in capabilities including, Task pane, Custom Function and Shortcut, Create Declarative Agent with Office Add-in Action, and Upgrade an Existing Office Add-in.":::

1. In the refined capability list that opens, select **New Declarative Agent with Office Add-in Actions**.

    :::image type="content" source="../images/agent-toolkit-refined-add-in-capability-list.png" alt-text="A dropdown list of two choices: New Declarative Agent with Office Add-in Action, and Extend an Existing Office Add-in.":::

1. In the **Workspace Folder** control that opens, choose a folder for the project.
1. In the **Application Name** text box, enter "Add-in + Agent Actions".
1. The project opens in a new Visual Studio Code window. Close the original Visual Studio Code window.
1. Open the Agents Toolkit, and in the **ACCOUNTS** section, ensure that you are logged into your Microsoft 365 tenancy and that both **Custom App Upload** and **Copilot Access** are enabled.

## Run the project

1. Close all Office applications.
1. Select **View | Run** in Visual Studio Code. In the **RUN AND DEBUG** dropdown menu, select **{{HOST}} Desktop (Edge Chromium)**, where {{HOST}} is `Excel`, `Powerpoint`, or `Word`. 

1. Press <kbd>F5</kbd>. The project builds and a Node dev-server window opens. This process may take a couple of minutes. Eventually, the Office application opens.

> [!NOTE]
> If this is the first time that you've sideloaded an Office Add-in on your computer (or the first time in over a month), you may be prompted to delete an old certificate and/or to install a new one. Agree to both prompts.

You can start working with either the add-in or the Copilot agent. If both have been started, then they each have an icon tab on the right side of the task pane so you can switch between them. 

### Start the add-in

1. There should be a **Contoso Add-in** group on the **Home** tab of the ribbon. If it isn't there, select the **Add-ins** button on the ribbon, and then select the **Add-in + Agent Actions** app in the flyout that opens.
1. The **Contoso Add-in** group has a **Show Taskpane** button which opens the task pane and a **Perform an action** button.

   > [!NOTE]
   > If a **WebView Stop On Load** prompt appears, select **OK**.

1. Test the add-in by pressing the **Run** link in the task pane, or selecting the **Perform an action** button.

### Start the agent

1. Select the floating icon for the **Copilot Chat entry point**.

    :::image type="content" source="../images/copilot-entry-point-floating-icon.png" alt-text="A round icon about the size of a button with the Copilot symbol. There is no text on the icon.":::

1. In the **Copilot** pane, select the hamburger control.
1. It may take a minute for the pane to completely rerender. When it does, there is a list of agents and **Add-in Skill + Agent for Add-in + Agent Actions** should be in the list. 
1. When the agent is listed, select it. The **Add-in Skill + Agent for Add-in + Agent Actions** pane opens.
1. Some conversation starters are listed. Select one that's appropriate for the Office application you have opened, and then press the **Send** control in the conversation box at the bottom of the pane. 
1. Select **Confirm** in response to the confirmation prompt.

   > [!TIP]
   > If Copilot reports an error, repeat your prompt but add the following sentence to the prompt: "If you get an error, report the complete text of the error to me."

1. Try entering other instructions in the conversation box. If you request an action that isn't defined in the agent, Copilot responds that it isn't able to do the action.

## Shut down the session completely

It's important to shut down the debugging session and uninstall the add-in and agent completely to avoid subtle problems. Use the following steps.

1. In Visual Studio Code, open the **Run** menu and select **Stop debugging**, or press <kbd>Shift</kbd>+<kbd>F5</kbd>. Due to a bug we're working on, this action doesn't always completely shut down the server, close the Office application, and uninstall the add-in. So, carry out the remaining steps.
1. Close the Office application if it's still running.
1. If the dev server is still running, shutting down the server depends on what window it's running in.

   - If the web server is running in a separate window from Visual Studio Code, open a command prompt or the Visual Studio Code **TERMINAL**. In the root of the project, run `npm run stop`.
   - If the web server is running in the Visual Studio Code **TERMINAL**, give the window focus and press <kbd>Ctrl</kbd>+<kbd>C</kbd>. Choose "Y" in response to the prompt to end the process. 

1. Clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
1. If the **Add-in + Agent Actions** is in the list of apps, select the arrow head to the left of the name to expand its row.
1. Select the trash can icon near the right end of the row, and then select **Remove** in the prompt.

## Troubleshooting

See [Troubleshooting combined add-ins and agents](../design/agent-and-add-in-overview.md#troubleshooting-combined-agents-and-add-ins).

## Next steps

1. Complete the tutorial that begins at [Create declarative agents using Microsoft 365 Agent Toolkit](/microsoft-365-copilot/extensibility/build-declarative-agents).
1. [Add a Copilot agent to an existing add-in](../develop/agent-and-add-in.md).

## See also

- [Combine Copilot Agents with Office Add-ins](../design/agent-and-add-in-overview.md)
- [Declarative agent manifest object](/microsoft-365-copilot/extensibility/declarative-agent-manifest-1.2#declarative-agent-manifest-object)
- [API plugin manifest schema 2.3 for Microsoft 365 Copilot](/microsoft-365-copilot/extensibility/api-plugin-manifest-2.3)