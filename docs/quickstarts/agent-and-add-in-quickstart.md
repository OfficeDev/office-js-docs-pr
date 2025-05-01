---
title: Make an Excel add-in into a Copilot skill
description: Learn how to build a simple Copilot agent that has an Excel add-in as a skill.
ms.date: 05/19/2025
ms.service: excel
ms.localizationpriority: high
---

# Build a Copilot agent with an Excel skill

In this article, you'll walk through the process of building a simple Excel Copilot agent that can perform some simple actions on the content of an Excel workbook. The app also includes an Excel task pane add-in.

## Knowledge prerequisites

- A basic understanding of declarative agents in Microsoft 365 Copilot. If you aren't familiar with them already, we recommend the following actions.
    - Read [Declarative agents for Microsoft 365 Copilot overview](/microsoft-365-copilot/extensibility/overview-declarative-agent).
    - Complete the tutorial that begins at [Create declarative agents using Teams Toolkit](/microsoft-365-copilot/extensibility/build-declarative-agents).

## Software prerequisites

- All the prerequisites listed at [Create declarative agents using Teams Toolkit](/microsoft-365-copilot/extensibility/build-declarative-agents).

## Start with an Office add-in

Begin by installing the prerelease version of Teams Toolkit. See [Install Teams Toolkit - Prerelease](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode#install-a-prerelease-version).

1. Create an Office Add-in in Teams Toolkit by following the instructions in [Create Office Add-in projects with Teams Toolkit](../develop/teams-toolkit-overview.md). *Stop after the project is created. Do not carry out the steps in the sideloading section.*

   > [!NOTE]
   > When prompted to name the add-in, use "Excel Add-in + Agent".

1. The project opens in a new Visual Studio Code window. Close the original Visual Studio Code window.
1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm install`.
1. The project is initially configured to support several Office applications. Remove non-Excel artifacts with the following steps:

   1. Open the manifest.json file in the appPackage folder of the project.
   1. The "authorization.permissions.resourceSpecific" array has two objects in it. Remove the object that mentions a Mailbox permission.
   1. In the extensions.requirements.scopes" array, remove all items except "workbook".
   1. The "extensions.runtimes" array has three objects. Remove first one, which is for Mailbox.
   1. The "extensions.ribbons" array has two objects. Remove first one, which is for Mailbox.
   1. In the remaining ribbon object, navigate to the "tabs.groups.controls" array. There are two control objects in the array. Remove the second one, which has name "ActionButton".

### Sideload and test the add-in

1. Test that the add-in works by carrying out the following steps.

   1. Select **View | Run** in Visual Studio Code. In the **RUN AND DEBUG** dropdown menu, select **Excel Desktop (Edge Chromium)**.
   1. Press F5. The project builds and a Node dev-server window opens. This process may take a couple of minutes. Eventually, Excel opens.

   > [!NOTE]
   > If this is the first time that you have sideloaded an Office Add-in on your computer (or the first time in over a month), you may be prompted to delete an old certificate and/or to install a new one. Agree to both prompts.

   1. Select the **Add-ins** button on the **Home** ribbon, and then in the flyout that opens, select your add-in. If you converted an existing Excel add-in, exercise its functions to verify that it works. The remaining steps in this section assume that you created a new add-in as instructed above in [Create an Office Add-in](#create-an-office-add-in).

   1. A **Contoso Add-in** group with a **Show Taskpane** button will appear on the **Home** ribbon. Use the button to open the add-in's task pane.

   > [!NOTE]
   > A **WebView Stop On Load** prompt appears. Select **OK**.

   1. When the task pane has opened, select **Run**. A cell in the worksheet changes to yellow.
   1. Stop debugging and uninstall the add-in by running `npm run stop` in a command prompt or Visual Studio Code **TERMINAL** in the root of the project.

      > [!IMPORTANT]
      > Stopping debugging in the UI of Visual Studio Code doesn't work currently due to a bug. Also, neither closing Excel nor manually closing the dev server window reliably shut down the server or cause Excel to unacquire the add-in. You **must** run `npm run stop`.

## Add a Copilot declarative agent

Add the agent with the following steps:

1. In the manifest file, make the following changes:

   1. Add the following object to the root. By convention, it is put just below the "validDomains" property. You create the "declarativeCopilot.json" file in a later step.

      ```json
      "copilotAgents": {
        "declarativeAgents": [
          {
            "id": "ContosoCopilotAgent",
            "file": "declarativeCopilot.json"
          }
        ]
      },
      ```

    1. In the second "extensions.runtimes" object, change the "actions.id" property to "fillcolor". This is the ID of a function that you add in a later step.
    1. In the same action object, change the "actions.type" property to "executeDataFunction".
    1. Change the "extensions.ribbons.tabs.groups.id" value to "BaseGroup".
    1. Change the "extensions.ribbons.tabs.groups.controls.id" value to "OpenTaskpane".

1. Create a file in the **appPackage** folder named "declarativeAgent.json".
1. Paste the following content into the file. You create the "Excel-API-local-plugin.json" file in a later step. For more information about these properties, see [Declarative agent manifest object](/microsoft-365-copilot/extensibility/declarative-agent-manifest-1.2#declarative-agent-manifest-object).

   ```json
   {
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

1. Create a file in the **appPackage** folder named "Excel-API-local-plugin.json".
1. Paste the following content into the file. For more information about these properties, see [API plugin manifest schema 2.2 for Microsoft 365 Copilot](/microsoft-365-copilot/extensibility/api-plugin-manifest-2.2). 

   > [!NOTE]
   > - The "runtimes.spec.local_endpoint" property is new. It tells the Copilot agent to look for functions in an add-in in Office instead of at a REST service URL.
   > - Any string in the "runtimes.run_for_functions" array must be an exact match for a "extensions.runtimes.actions.id" property in the manifest.

   ```json
   {
        "schema_version": "v2.2",
        "name_for_human": "Excel Add-in + Agent",
        "description_for_human": "Add-in Actions in Agents",
        "namespace": "add-in_function",
        "functions": [
            {
                "name": "fillcolor",
                "description": "fillcolor changes a single cell location to a specific color.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "Cell": {
                            "type": "string",
                            "description": "A cell location in the format of A1, B2,
                            etc.",
                            "default" : "B2"
                        },
                        "Color": {
                            "type": "string",
                            "description": "A color in hex format, e.g. #30d5c8",
                            "default" : "#30d5c8"
                        }
                    },
                    "required": ["Cell", "Color"]
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
                    "local_endpoint": "ms-office-addin"
                },
                "run_for_functions": ["fillcolor"]
            }
        ]
    }
   ```

1. Open the \src\commands.ts file and replace its contents with the following:

   ```javascript
   async function fillColor(cell, color) {
       await Excel.run(async (context) => {
            context.workbook.worksheets
                .getActiveWorksheet()
                .getRange(cell).format.fill.color = color;
            await context.sync();
        })
   }

   Office.onReady((info) => {
      Office.actions.associate("fillcolor", async (message) => {
        const {Cell: cell, Color: color} = JSON.parse(message);
        await fillColor(cell, color);
        return result;
      });
   });
   ```

## Update project configuration files for a combined add-in and Copilot agent

1. Open the teamsapp.yaml file in the root of project and replace its contents with the following.

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
   # Apply the Teams app manifest to an existing Teams app in
   # Teams Developer Portal.
   # Will use the app id in manifest file to determine which Teams app to update.
   - uses: teamsApp/update
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
1. Open the \env\.env.dev file and add the following lines to the end of the file, right after the line "ADDIN_ENDPOINT=".

   ```
   TEAMS_APP_ID=
   TEAMS_APP_TENANT_ID=
   M365_TITLE_ID=
   M365_APP_ID=
   ```

## Test the add-in and agent

1. Close all Office applications.
1. Open Teams Toolkit.
1. In the **Lifecycle** pane, select **Provision**. Among other things, provisioning will do the following:

   - Set values for the four lines you added to the .env.dev file.
   - Create a /build folder inside the /appPackage folder with the package zip file. The file contains the manifest and JSON files for the agent and plugin.

1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
1. Select **Upload an app** in the **Apps** dialog, and then in the dialog that opens, select **Upload a custom app**.
1. In the **Open** dialog, select the package zip file in the project's /appPackage/build folder.
1. Select **Add** in the dialog that opens.
1. When you are prompted that the app was added, *don't* open it in Teams. Instead close Teams.
1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run dev-server` to start the server on localhost.
1. Open Excel. In a few moments, the **Show Task pane** button appears on the **Home** ribbon in the **Contoso Add-in** group.
1. Open **Copilot** from the ribbon and select the hamburger control in the **Copilot** pane. **Excel Add-in + Agent** should be be in the list of agents. If it is not, wait a few minutes and reload Copilot.
1. When the agent is listed, select it and the  **Excel Add-in + Agent** pane opens.
1. Select the **Change cell color** conversation starter, and then press the **Send** control in the conversation box at the bottom of the pane. Select **Confirm** in response to the confirmation prompt. The cell's color is changed.
1. Try entering other combinations of cell and color in the conversation box, such as "Set cell G5 to the color of the sky".

## Make changes in the add-in or agent

Live reloading and hot reloading for a combined add-in and agent aren't supported in the preview period. To make changes, first shut down the server and uninstall the extension with these steps.

1. Close Excel.
1. If the web server is running in the Visual Studio Code **TERMINAL**, give the terminal focus and press Ctrl-C. Choose "Y" in response to the prompt to end the process. Then go to the next step. If the web server is running in a separate window skip this step and go to the next step.
1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run stop`.
1. Clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
1. Find "Excel Add-in + Agent" in the list of apps, and select the arrow head to the left of the name to expand its row.
1. Select the trash can icon near the right end of the row, and then select **Remove** in the prompt.

Make your changes and then repeat the steps in [Test the add-in and agent](#test-the-add-in-and-agent).
