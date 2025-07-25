---
title: Build your first add-in as a Copilot skill
description: Learn how to build a simple Copilot agent that has an Excel add-in as a skill.
ms.date: 07/24/2025
ms.topic: how-to
ms.service: microsoft-365
ms.localizationpriority: high
---

# Build your first add-in as a Copilot skill

In this article, you'll walk through the process of building a simple Excel Copilot agent that can perform actions on the content of an Excel workbook. The app also includes an Excel task pane add-in.

## Knowledge prerequisites

- A basic understanding of declarative agents in Microsoft 365 Copilot. If you aren't familiar with them already, we recommend the following actions.
    - Read [Declarative agents for Microsoft 365 Copilot overview](/microsoft-365-copilot/extensibility/overview-declarative-agent).
    - Complete the tutorial that begins at [Create declarative agents using Microsoft 365 Agent Toolkit](/microsoft-365-copilot/extensibility/build-declarative-agents).

## Software prerequisites

- All the prerequisites listed at [Create declarative agents using Microsoft 365 Agent Toolkit](/microsoft-365-copilot/extensibility/build-declarative-agents).
- The [Microsoft 365 Agent Toolkit](../develop/teams-toolkit-overview.md).

## Start with an Office Add-in

Create a basic Excel add-in with the following steps.

1. Create an Office Add-in in Microsoft 365 Agent Toolkit by following the instructions in [Create Office Add-in projects with Microsoft 365 Agent Toolkit](../develop/teams-toolkit-overview.md#create-an-office-add-in-project). *Stop after the project is created. Don't carry out the steps in the sideloading section.*

   > [!NOTE]
   > When prompted to name the add-in, use "Excel Add-in + Agent".

1. The project opens in a new Visual Studio Code window. Close the original Visual Studio Code window.
1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm install`.

### Sideload and test the add-in

1. Test that the add-in works by carrying out the following steps.

   1. Select **View | Run** in Visual Studio Code. In the **RUN AND DEBUG** dropdown menu, select **Excel Desktop (Edge Chromium)**.
   1. Press <kbd>F5</kbd>. The project builds and a Node dev-server window opens. This process may take a couple of minutes. Eventually, Excel opens.

   > [!NOTE]
   > If this is the first time that you have sideloaded an Office Add-in on your computer (or the first time in over a month), you may be prompted to delete an old certificate and/or to install a new one. Agree to both prompts.

   1. Select the **Add-ins** button on the **Home** ribbon, and then in the flyout that opens, select your add-in.
   1. A **Contoso Add-in** group with a **Show Taskpane** button will appear on the **Home** ribbon. Use the button to open the add-in's task pane.

   > [!NOTE]
   > If a **WebView Stop On Load** prompt appears, select **OK**.

   1. When the task pane has opened, select **Run**. A cell in the worksheet changes to yellow.
   1. Stop debugging and uninstall the add-in by shutting down Excel and running `npm run stop` in a command prompt or Visual Studio Code **TERMINAL** in the root of the project.

      > [!IMPORTANT]
      > Stopping debugging in the UI of Visual Studio Code doesn't work currently due to a bug. Also, neither closing Excel nor manually closing the dev server window reliably shut down the server or cause Excel to unacquire the add-in. You **must** run `npm run stop`.

## Add a Copilot declarative agent

Add the agent with the following steps.

1. In the manifest file, make the following changes.

   1. Add the following object to the root. By convention, it's put just below the "validDomains" property. You create the "declarativeAgent.json" file in a later step.

      ```json
      "copilotAgents": {
        "declarativeAgents": [
          {
            "id": "ContosoCopilotAgent",
            "file": "declarativeAgent.json"
          }
        ]
      },
      ```

    1. There are multiple objects in the `"extensions.runtimes"` array. Find the one whose `"id"` is "CommandRuntime" and copy it as an additional runtime object in the array.
    1. Make the following changes to this additional runtime object. 
    
       1. Change the `"id"` from "CommandRuntime" to "CopilotAgentActionsRuntime".
       1. Change its `"actions.id"` property to "fillcolor". This is the ID of a function that you add in a later step.
       1. Change the `"actions.type"` property to "executeDataFunction".

1. Create a file in the **appPackage** folder named **declarativeAgent.json**.
1. Paste the following content into the file. (You create the **Excel-API-local-plugin.json** file that is mentioned in this JSON in a later step.)

   ```json
   {
        "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.4/schema.json",
        "version": "v1.4",
        "name": "Excel Add-in + Agent",
        "description": "Agent for working with Excel cells.",
        "instructions": "You are an agent for working with an add-in. You can work with any cells, not just a well-formatted table.",
        "conversation_starters": [
            {
                "title": "Change cell color",
                "text": "I want to change the color of cell B2 to orange"
            }
        ],
        "actions": [
            {
                "id": "localExcelPlugin",
                "file": "Excel-API-local-plugin.json"
            }
        ]
    }
   ```

1. Create a file in the **appPackage** folder named **Excel-API-local-plugin.json**.
1. Paste the following content into the file.

   ```json
   {
        "$schema": "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.3/schema.json",
        "schema_version": "v2.3",
        "name_for_human": "Excel Add-in + Agent",
        "description_for_human": "Add-in Actions in Agents",
        "namespace": "addinfunction",
        "functions": [
            {
                "name": "fillcolor",
                "description": "fillcolor changes a single cell location to a specific color.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "cell": {
                            "type": "string",
                            "description": "A cell location in the format of A1, B2, etc.",
                            "default" : "B2"
                        },
                        "color": {
                            "type": "string",
                            "description": "A color in hex format, e.g., #30d5c8",
                            "default" : "#30d5c8"
                        }
                    },
                    "required": ["cell", "color"]
                },
                "returns": {
                    "type": "string",
                    "description": "A string indicating the result of the action."
                },
                "states": {
                    "reasoning": {
                        "description": "`fillcolor` changes the color of a single cell based on the grid location and a color value.",
                        "instructions": "The user will pass ask for a color that isn't in the hex format needed in most cases, make sure to convert to the closest approximation in the right format."
                    },
                    "responding": {
                        "description": "`fillcolor` changes the color of a single cell based on the grid location and a color value.",
                        "instructions": "If there is no error present, tell the user the cell location and color that was set."
                    }
                }
            }
        ],
        "runtimes": [
            {
                "type": "LocalPlugin",
                "spec": {
                    "local_endpoint": "Microsoft.Office.Addin",
                    "allowed_host": ["workbook"]
                },
                "run_for_functions": ["fillcolor"]
            }
        ]
    }
   ```

1. Open the **\src\commands\commands.ts** file and add the following code the end of it.

   ```javascript
   async function fillcolor(cell, color) {
       await Excel.run(async (context) => {
            context.workbook.worksheets
                .getActiveWorksheet()
                .getRange(cell).format.fill.color = color;
            await context.sync();
        })
   }

   Office.onReady((info) => {
        Office.actions.associate("fillcolor", async (message) => {
            const {cell, color} = JSON.parse(message);
            await fillcolor(cell, color);
            return "Cell color changed.";
        });
   });
   ```

## Update project configuration files for a combined add-in and Copilot agent

1. There is a file called either **teamsapp.yaml** or **m365agents.yaml** in the root of project. Replace its contents with the following:

   ```yaml
   # yaml-language-server: $schema=https://aka.ms/teams-toolkit/v1.7/yaml.schema.json
   # Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
   # Visit https://aka.ms/teamsfx-actions for details on actions
   version: v1.7

   environmentFolderPath: ./env

   # Triggered when 'teamsapp provision' is executed
   provision:
   # Creates a Teams app
     - uses: teamsApp/create
       with:
        # Teams app name
        name: Contoso Agent ${{APP_NAME_SUFFIX}}
       # Write the information of created resources into environment file for
       # the specified environment variable(s).
       writeToEnvironmentFile:
        teamsAppId: TEAMS_APP_ID

   # Build Teams app package with latest env value
     - uses: teamsApp/zipAppPackage
       with:
        # Path to manifest template
        manifestPath: ./appPackage/manifest.json
        outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
        outputFolder: ./appPackage/build
   # Validate app package using validation rules
     - uses: teamsApp/validateAppPackage
       with:
        # Relative path to this file. This is the path for built zip file.
        appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
   # Extend your Teams app to Outlook and the Microsoft 365 app
     - uses: teamsApp/extendToM365
       with:
        # Relative path to the build app package.
        appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
       # Write the information of created resources into environment file for
       # the specified environment variable(s).
       writeToEnvironmentFile:
        titleId: M365_TITLE_ID
        appId: M365_APP_ID

   # Triggered when 'teamsapp publish' is executed
   publish:
   # Build Teams app package with latest env value
     - uses: teamsApp/zipAppPackage
       with:
        # Path to manifest template
        manifestPath: ./appPackage/manifest.json
        outputZipPath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
        outputFolder: ./appPackage/build
   # Validate app package using validation rules
     - uses: teamsApp/validateAppPackage
       with:
        # Relative path to this file. This is the path for built zip file.
        appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
   # Apply the Teams app manifest to an existing Teams app in
   # Teams Developer Portal.
   # Will use the app id in manifest file to determine which Teams app to update.
     - uses: teamsApp/update
       with:
        # Relative path to this file. This is the path for built zip file.
        appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
   # Publish the app to
   # Teams Admin Center (https://admin.teams.microsoft.com/policies/manage-apps)
   # for review and approval
     - uses: teamsApp/publishAppPackage
       with:
        appPackagePath: ./appPackage/build/appPackage.${{TEAMSFX_ENV}}.zip
       # Write the information of created resources into environment file for
       # the specified environment variable(s).
       writeToEnvironmentFile:
        publishedAppId: TEAMS_APP_PUBLISHED_APP_ID
   projectId: da53b0a2-1561-415e-919a-5b870bcd2f49
   ```

1. Replace the value of `projectId` in the last line of content you pasted in the preceding step with a new randomly generated GUID.
1. Open the **\env\.env.dev** file and add the following lines to the end of the file, right after the line "ADDIN_ENDPOINT=".

   ```
   TEAMS_APP_ID=
   TEAMS_APP_TENANT_ID=
   M365_TITLE_ID=
   M365_APP_ID=
   ```

## Test the add-in and agent

1. Close all Office applications.
1. Open Microsoft 365 Agent Toolkit.
1. In the **Lifecycle** pane, select **Provision**. Among other things, provisioning does the following:

   - Set values for the four lines you added to the .env.dev file.
   - Create a **/build** folder inside the **/appPackage** folder with the package zip file. The file contains the manifest and JSON files for the agent and plug-in.

1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run dev-server` to start the server on localhost. Wait until you see a line in the server window that the app compiled successfully. This means the server is running and serving the files.
    
    > [!NOTE]
    > If this is the first time in over a month you have run a local server for an Office Add-in on your computer, you may be prompted to delete an old certificate and to install a new one. Agree to both prompts.

1. The first step in testing depends on the platform.

   - To test in Office on Windows, open Excel. In a few moments, the **Show Task pane** button appears on the **Home** ribbon in the Contoso Add-in group. (If it doesn't appear on the ribbon, select the **Add-ins** button on the ribbon, and then select the **Excel Add-in + Agent** app in the flyout that opens.)
   - To test in Office on the web, in a browser, navigate to `https://excel.cloud.microsoft.com/`, and then create a new workbook.
 
1. The process of opening your agent depends on the UI for Copilot in Office applications which is in transition.

   - If there is a **Copilot** *button* on the ribbon (not a dro down menu), select the **Copilot** button to open the **Copilot** pane.
   - If there is a **Copilot** dropdown menu, open the menu and select **App Skills** to open the **Copilot** pane.

1. In the **Copilot** pane, select the hamburger control.
1. In the pane, **Excel Add-in + Agent** should be in the list of agents. (You may need to select **See more** to ensure that all agents are listed.) If the agent isn't, try one or both of the following actions.

   - Wait a few minutes and reload Copilot.
   - With Copilot open to the list of agents, click the cursor on the **Copilot** pane and press <kbd>Ctrl</kbd>+<kbd>R</kbd>.

1. When the agent is listed, select it. The **Excel Add-in + Agent** pane opens.
1. Select the **Change cell color** conversation starter, and then press the **Send** control in the conversation box at the bottom of the pane. Select **Confirm** in response to the confirmation prompt. The cell's color should change.

   > [!TIP]
   > If Copilot reports an error, repeat your prompt but add the following sentence to the prompt: "If you get an error, report the complete text of the error to me."

1. Try entering other combinations of cell and color in the conversation box, such as "Set cell G5 to the color of the sky".

## Make changes in the add-in or agent

Live reloading and hot reloading for a combined add-in and agent aren't supported in the preview period. To make changes, first shut down the server and uninstall the extension with these steps.

1. Shutting down the server depends on what window it's running in.

   - If the web server is running in the same command prompt or Visual Studio Code **TERMINAL** where you ran `npm run dev-server`, give the window focus and press <kbd>Ctrl</kbd>+<kbd>C</kbd>. Choose "Y" in response to the prompt to end the process. 
   - If the web server is running in a separate window, then in a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run stop`.

1. Clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
1. Find **Excel Add-in + Agent** in the list of apps, and select the arrow head to the left of the name to expand its row.
1. Select the trash can icon near the right end of the row, and then select **Remove** in the prompt.

Make your changes and then repeat the steps in [Test the add-in and agent](#test-the-add-in-and-agent).

## Troubleshooting

See [Troubleshooting combined add-ins and agents](../design/agent-and-add-in-overview.md#troubleshooting-combined-agents-and-add-ins).

## Next steps

- [Add a Copilot agent to an existing add-in](../develop/agent-and-add-in.md)

## See also

- [Combine Copilot Agents with Office Add-ins](../design/agent-and-add-in-overview.md)
- [Declarative agent manifest object](/microsoft-365-copilot/extensibility/declarative-agent-manifest-1.2#declarative-agent-manifest-object)
- [API plugin manifest schema 2.3 for Microsoft 365 Copilot](/microsoft-365-copilot/extensibility/api-plugin-manifest-2.3)