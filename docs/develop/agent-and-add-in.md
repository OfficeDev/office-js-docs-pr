---
title: Add a Copilot agent to an add-in
description: Learn how to add a Copilot agent to an add-in.
ms.date: 10/17/2025
ms.topic: how-to
ms.service: microsoft-365
ms.localizationpriority: medium
---

# Add a Copilot agent to an add-in

Adding a Copilot agent to an Office Add-in provides two benefits:

- Copilot becomes a natural language interface for the add-in's functionality.
- The agent can pass parameters to the JavaScript it invokes, which isn't possible when a [function command](../design/add-in-commands.md#types-of-add-in-commands) is invoked from a button or menu item.

> [!NOTE]
> This article assumes that you're familiar with the overview [Combine Copilot Agents with Office Add-ins](../design/agent-and-add-in-overview.md) and the Copilot documentation that it refers to. We also recommend that you complete the quick start [Build your first add-in as a Copilot skill](../quickstarts/agent-and-add-in-quickstart.md).

> [!IMPORTANT]
> This feature requires the [unified manifest for Microsoft 365](unified-manifest-overview.md). If your add-in uses the add-in only manifest, you must first [convert it to use the unified manifest](convert-xml-to-json-manifest.md) before you can add a Copilot agent to it. Before you continue with this article, you should know how to package the manifest and other files into an app package zip file and sideload it to an Office application for testing.

## Major tasks

The following are the main tasks for adding a Copilot agent to your add-in. Details are in the subsections.

- [Create the functions for the agent's actions](#create-the-functions-for-the-agents-actions) that will implement the Copilot agent's actions.
- [Update the manifest](#update-the-manifest)
- [Create the agent and API plug-in configuration](#create-the-agent-and-api-plug-in-configuration)
- [Create the app package](#create-the-app-package)
- [Test the agent](#test-the-agent)
- [Make changes in the app](#make-changes-in-the-app)

### Create the functions for the agent's actions

If your add-in includes one or more [function commands](../design/add-in-commands.md#types-of-add-in-commands), then your project already has a JavaScript or TypeScript file that defines the functions for these commands (usually called **commands.js** or **commands.ts**) and a UI-less HTML file (usually called **commands.html**) that has a `<script>` tag to load the function file. We recommend that you use this same function file to define the functions for your agent's actions. Skip to the section [Update the function file](#update-the-function-file).

#### Create the source files

If your project doesn't already have such files, then create them with the following steps. The folder and file structure and names in these steps aren't mandatory, but we recommend them to minimize incompatibilities with add-in development tooling and configuration for bundlers and transpilers where applicable.

1. Ensure that there's a folder **\src** off the root of the project, and that it has a child folder named **\commands**.
1. Create a **commands.js** (or .ts) file in the **\commands** folder. You add content to it in a later step.
1. Create a **commands.html** file in the **\commands** folder with the following content. Note the following about this markup.

   - The `<body>` element is empty because the file has no UI. Its only purpose is to load JavaScript files.
   - The Office JavaScript Library and the **commands.js** file that you created in the preceding step is explicitly loaded.

       > [!NOTE]
       > It's common in Office Add-in development to use tools like [webpack](https://webpack.js.org/) and its plug-ins to automatically inject `<script>` tags into HTML files at build time. If you use such a tool, you shouldn't include any `<script>` tags in your source file that are going to be inserted by the tool.

   ```html
   <!DOCTYPE html>
   <html>
       <head>
           <meta charset="UTF-8" />
           <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

           <!-- Office JavaScript Library -->
           <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
           <!-- Function command file -->
           <script src="commands.js" type="text/javascript"></script>
       </head>
       <body>
       </body>
   </html>
   ```

#### Update the function file

Define the functions for your agent's actions with the following steps. These assume the most common folder and file structure and names.

Open the function file (usually, the **\src\commands\commands.js** (or .ts)) and add the functions that implement the agent's actions. Keep the following points in mind as you work.

- Like all code that calls APIs in the Office JavaScript Library, the file must [initialize the library](initialize-add-in.md), usually by calling `Office.onReady`.
- Functions that are invoked by agents take a `message` parameter that specifies the values that Office will use when it calls functions from the Office JavaScript Library. This object can be parsed with the `JSON.parse` method. The values of the object are specified by users in natural language.
- Unlike the functions that are invoked with add-in commands, these functions don't have an `Office.AddinCommands.Event` parameter, and they don't call the `Office.AddinCommands.Event.completed` method. Instead, functions that are invoked by agents return a result object from the JavaScript runtime back to the Copilot runtime.
- Functions invoked from Copilot have less time to complete than functions invoked from a function command. The latter have five minutes to complete or Office shuts down the JavaScript runtime. But Copilot-invoked functions have only two minutes to return a result or the runtime is shut down.
- For each function, there must be a call of [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) to tell Office which function in the file should be run when the agent action is invoked. The `associate` function maps the function name to an action ID that you configure in the manifest in a later step. If you define multiple functions in the same file, your code must call `associate` for each one.

The following is an example.

```javascript
Office.onReady(function() {
    // Add any initialization code here.
});

async function fillColorFromUserData(message) {
    const {cell, color} = JSON.parse(message);
    await Excel.run(async (context) => {
      context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(cell).format.fill.color = color;
      await context.sync();
    });
    return "Cell color changed.";
}

Office.actions.associate("FillColor", fillColorFromUserData);
```

### Update the manifest

There are three major parts to configuring the manifest for the Copilot agent as described in the following subsections.

#### Use the preview manifest schema

Ensure that the manifest references the preview version of the manifest schema with the following steps.

1. At the top of the manifest, ensure that the [`"$schema"`](/microsoft-365/extensibility/schema/root#schema-4) property is the following:

   ```json
   "$schema": "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
   ```

1. Ensure that the [`"manifestVersion"`](/microsoft-365/extensibility/schema/root#manifestversion) property is set to "devPreview".

#### Configure the runtime

1. In the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array) array, ensure that there's a runtime object that is oriented to running UI-less JavaScript functions. It's critical that the object have a [`"code.page"`](/microsoft-365/extensibility/schema/extension-runtime-code#page) property that specifies the absolute URL to the UI-less HTML file that you created (or edited) earlier in [Update the function file](#update-the-function-file).
1. In the [`"actions"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item) array of that same runtime object, add an object whose `"type"` property is set to "executeDataFunction" and whose `"id"` property is an exact match for the first parameter of the call to `Office.actions.associate` in a function that you created in [Update the function file](#update-the-function-file).
1. Repeat the preceding step for each function you created to implement an agent action.

The runtime object should look similar to the following. There may be other properties of the runtime object, other runtime objects in the `"runtimes"` array, and other action objects in the `"actions"` array.

```json
"runtimes": [
    {
        "id": "CommandsRuntime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "FillColor",
                "type": "executeDataFunction",
            }
        ]
    }
]
```

#### Declare the agent

1. If there isn't one already, add a ["copilotAgents"](/microsoft-365/extensibility/schema/root-copilot-agents) object to the root of the manifest and ensure that it has a child ["declarativeAgents"](/microsoft-365/extensibility/schema/declarative-agent-ref) array.
1. Add an object to the `"declarativeAgents"` array and specify a unique and descriptive [`"id"`](/microsoft-365/extensibility/schema/declarative-agent-ref#id) for it.
1. Assign the object's [`"file"`](/microsoft-365/extensibility/schema/declarative-agent-ref#file) property the relative URL of the declarative agent configuration file. You create that file in a later step.

   The following is an example.

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

> [!TIP]
> To maximize compatibility with Microsoft tools for add-in development, we recommend that you create a folder named **appPackage** in the root of the project and move the manifest into it.

### Create the agent and API plug-in configuration

1. Create a file in the same folder where your manifest is and give it the name used in the `"copilot.declarativeAgents.file"` property, such as **declarativeAgent.json**.
1. Paste the following content into the file.

   ```json
   {
        "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.5/schema.json",
        "version": "v1.5",
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
                "file": "Office-API-local-plugin.json"
            }
        ]
    }
   ```

    [!INCLUDE [Validation warning about missing 'auth' property](../includes/auth-property-warning-note.md)]

1. *Replace the property values with new values that are appropriate for your add-in.* For more information about these properties, see [Declarative agent manifest object](/microsoft-365-copilot/extensibility/declarative-agent-manifest-1.2#declarative-agent-manifest-object).

   > [!NOTE]
   > You create the file that you specify in the `"actions.file"` property in the next step. In the example above, this is the file **Office-API-local-plugin.json**.

1. Create another file in the folder with your manifest and give it the name you assigned to the `"actions.file"` property in the preceding step; for example, **Office-API-local-plugin.json**.
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
                "name": "FillColor",
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
                        "description": "`FillColor` changes the color of a single cell based on the grid location and a color value.",
                        "instructions": "The user will pass ask for a color that isn't in the hex format needed in most cases, make sure to convert to the closest approximation in the right format."
                    },
                    "responding": {
                        "description": "`FillColor` changes the color of a single cell based on the grid location and a color value.",
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
                "run_for_functions": ["FillColor"]
            }
        ]
    }
   ```

1. *With some exceptions noted later, replace the property values with new values that are appropriate for your add-in.* For more information about these properties, see [API plugin manifest schema 2.3 for Microsoft 365 Copilot](/microsoft-365-copilot/extensibility/api-plugin-manifest-2.3).

   As you work, keep the following points in mind.

   - Do *not* change the `"namespace"`, `"runtimes.type"`, or `"runtimes.spec.local_endpoint"` values.
   - The `"functions.name"` must be an exact match for both of the following:

      - An `"extensions.runtimes.actions.id"` property in the manifest (for an action of type "executeDataFunction").
      - The first parameter of the call to `Office.actions.associate` in a function that you created in [Update the function file](#update-the-function-file).

   - The `"runtimes.run_for_functions"` array must include either the same string as `"functions.name"` or a wildcard string that matches it.
   - The `"reasoning.description"` and `"reasoning.instructions"` refer to a JavaScript function, not a REST API.
   - The `"runtimes.spec.local_endpoint"` property tells the Copilot agent to look for functions in an Office Add-in instead of at a REST service URL.

### Create the app package

> [!IMPORTANT]
> To test the agent, as described later in this article in [Test the agent](#test-the-agent), the domain segment of any absolute URLs in the manifest must be a localhost domain; for example, `localhost:3000`. You can change these segments to a production domain later.

Using any zip utility, create a zip file that contains the following files.

- The manifest
- The icon files referenced in the manifest's [`"icons.color"`](/microsoft-365/extensibility/schema/root-icons#color) and [`"icons.outline"`](/microsoft-365/extensibility/schema/root-icons#outline)
- The two files you created in [Create the agent and API plug-in configuration](#create-the-agent-and-api-plug-in-configuration)
- Any files referenced in the [`"localizationInfo.additionalLanguages"`](/microsoft-365/extensibility/schema/root-localization-info#additionallanguages) property

> [!IMPORTANT]
> Most of these files have URLs in the manifest that are relative to the location of the manifest, so the folder structure of the zip file must maintain this structure. For example, if the `"icons.color"` value is "/assets/icon-32.png", then there must be an **/assets** folder in the zip file with the **icon-32.png** file in it.

> [!TIP]
> To maximize compatibility with Microsoft tools for add-in development, we recommend that you create a subfolder named **build** in the **appPackage** folder and create the zip file in it.

### Test the agent

The three major steps to test the agent &mdash; sideloading, starting a server, and running the agent &mdash; are described in the following subsections.

#### Sideload the agent and add-in

Sideloading is done through Teams even when there's no Teams feature in the app. The steps are as follows.

1. Close all Office applications, and then clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
1. Select **Upload an app** in the **Apps** dialog, and then in the dialog that opens, select **Upload a custom app**.
1. In the **Open** dialog, navigate to, and select, the package zip file.
1. Select **Add** in the dialog that opens.
1. When you're prompted that the app was added, *don't* open it in Teams. Instead, close Teams.

#### Start the server

Your task is to start a local web server that hosts your project's HTML and JavaScript files. How you do this depends on several factors including the folder structure of your project, the tools you use, such as a bundler, task manager, server application, and how you have configured those tools. Since you're adding an agent to an existing add-in project, you already know how to do this. The following instruction applies only to projects that meet the following conditions.

- There's a **webpack.config.js** file in the root of the project that is similar to the ones in add-in projects that are created with the [Yeoman Generator for Office Add-ins](yeoman-generator-overview.md) or [Microsoft 365 Agent Toolkit](teams-toolkit-overview.md).
- There's a **package.json** file in the root of the project similar to the ones created by the same two tools and the file has a "scripts" section with the following script in it.

   ```json
   "dev-server": "webpack serve --mode development"
   ```

In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run dev-server` to start the server on localhost.

#### Run the agent

1. Open the Office application (Excel, PowerPoint, or Word) that your combined agent and add-in targets. Wait until the add-in has loaded. This may take as much as two minutes. Depending on your version of Office, ribbon buttons and other artifacts may appear automatically. In recent versions, you need to manually activate the add-in: Select the **Add-ins** button on the **Home** ribbon, and then in the flyout that opens, select your add-in. It will have the name from the name from the [`"name.short"`](/microsoft-365/extensibility/schema/root-name) property in the manifest.
1. The process of opening your agent depends on the UI for Copilot in Office applications which is in transition.

   - If there is a **Copilot** *button* on the ribbon (not a dropdown menu), select the **Copilot** button to open the **Copilot** pane.
   - If there is a **Copilot** dropdown menu, open the menu and select **App Skills** to open the **Copilot** pane.

1. In the **Copilot** pane, select the hamburger control.
1. In the pane, your agent should be in the list of agents. It has the name specified in the `"name"` property of the declarative agent configuration file (which may not be the same as the name from the `"name.short"` property in the manifest); for example, **Excel Agent**. You may need to select **See more** to ensure that all agents are listed. If the agent isn't listed, try one or both of the following actions.

   - Wait a few minutes and reload Copilot.
   - With Copilot open to the list of agents, click the cursor on the **Copilot** pane and press <kbd>Ctrl</kbd>+<kbd>R</kbd>.

   :::image type="content" source="../images/copilot-agent-list.png" alt-text="A screenshot of the agent list in the Copilot pane in an Office application":::

1. When the agent is listed, select it and the pane for the agent opens. The conversation starters you configured in the `"conversation_starters"` property of declarative agent configuration file will be displayed.
1. Select a conversation starter, and then press the **Send** control in the conversation box at the bottom of the pane. Select **Confirm** in response to the confirmation prompt. The agent action occurs.
1. Try entering prompts the conversation box that are different from the conversation starters, but that your agent should be able to do.

### Make changes in the app

Live reloading and hot reloading for a combined add-in and agent aren't supported in the preview period. To make changes, first shut down the server and uninstall the agent and add-in with these steps.

1. Close the Office application.
1. If the web server is running in the Visual Studio Code **TERMINAL**, give the terminal focus and press <kbd>Ctrl</kbd>+<kbd>C</kbd>. Choose "Y" in response to the prompt to end the process. Then go to the next step. If the web server is running in a separate window, skip this step and go to the next step.
1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run stop`.
1. Clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
1. Find your agent in the list of apps. It will have the name specified in the `"name"` property of the declarative agent configuration file; for example, **Excel Add-in + Agent**.
1. Select the arrow head to the left of the name to expand its row.
1. Select the trash can icon near the right end of the row, and then select **Remove** in the prompt.

Make your changes and then repeat the steps in [Test the agent](#test-the-agent).

## Troubleshooting

See [Troubleshooting combined add-ins and agents](../design/agent-and-add-in-overview.md#troubleshooting-combined-agents-and-add-ins).
