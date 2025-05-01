---
title: Combine Copilot Agents with Office Add-ins (preview)
description: Get an overview of why and how to combine a Copilot agent with an Office Add-in.
ms.date: 05/19/2025
ms.topic: overview
ms.localizationpriority: medium
---

# Combine Copilot Agents with Office Add-ins (preview)

> [!NOTE]
> This article assumes you are familiar with Copilot declarative agents. If you're not, start with the following:
>
> - [Declarative agents for Microsoft 365 Copilot overview](/microsoft-365-copilot/extensibility/overview-declarative-agent).
> - [Agents are apps for Microsoft 365](/microsoft-365-copilot/extensibility/agents-are-apps).

Including a Microsoft 365 Copilot agent in an Office Add-in provides two benefits:

- Copilot becomes a natural language interface for the add-in's functionality.
- The agent can pass parameters to the JavaScript it invokes, which isn't possible when a [function command](add-in-commands.md#types-of-add-in-commands) is invoked from a button or menu item.

You can also think of an Office Add-in as a skill in a Copilot agent. Because Office add-ins use the [Office JavaScript Library](../develop/understanding-the-javascript-api-for-office.md) to perform read and write operations on Office documents, these operations become actions in the Copilot agent.

## Scenarios

The following are some selected ways in which adding a Copilot agent enhances the value of an add-in to users.

- **Learning how to use the add-in:** When a user needs to perform multiple steps or tasks with the add-in to achieve a goal, the chat interface of Copilot can ease the process of getting started with the add-in. For example, consider a legal firm that needs to have a list of questions that must be answered about each lease that it prepares. Creating this list of questions can be time-consuming and labor-intensive. But a Copilot agent that uses the Office JavaScript Library can be asked to produce a first draft list of questions and insert them into a Word document.

- **Content analysis:** An agent can be used to analyze the content of a document or spreadsheet and take action depending on what it finds. The following are examples.

   - An agent analyzes a Request for Proposal and then fetches the answers to questions in the RFP from a backend system. The user simply instructs the agent to "Fill in the answers you know to the questions."
   - An agent analyzes a document, or a table in a spreadsheet, for content that implies certain actions must be taken, either in the document itself or elsewhere in the customer's business systems. The user might say "Review the document for any items I missed on the audit list."

- **Trusted insertion of data:** If you ask a typical AI engine a question, it will combine information it finds and compose an answer; a process can introduce inaccuracies. But a Copilot agent based on an add-in can insert data *unchanged* from a trusted source. Some examples:

   - Consider an add-in that enables the insertion of legal research into Word where it can then be edited. A user asks the agent "In what circumstances can a lease of residential space in Indiana be broken unilaterally by the lessor?" The add-in then fetches content, unchanged, from precedents and statutes.
   - Consider an add-in that manages the inventory of a digital assets. In the Copilot agent chat, a user asks "Insert a table of our color photos with the name of each, the number of times it was downloaded, and it's size in megabytes, sorted in order from most downloaded." The add-in then fetches this data, unchanged, from the system of record and inserts the table into an Excel spreadsheet.

## The relation of Copilot agents to the Add-in framework

A Copilot agent is a natural language interface for an add-in.

An add-in can be configured to be *only* a skill in a Copilot agent. It doesn't have to include a task pane, custom ribbon buttons, or custom menus; but it can have any of these in addition to being a Copilot skill. The best approach depends on the user scenarios that the add-in should enable.

- If the add-in should provide simple, fast actions that do not need parameters passed to them, use custom ribbon buttons or menus, called [add-in commands](add-in-commands.md).
- If the add-in needs a dashboard experience, needs configuration of settings by the user, needs to display metadata about the content of the Office document, or needs a page-like experience for any other reason, include a task pane in the add-in.
- If the add-in needs to provide complex actions that require parameters passed at runtime or needs a natural language interface, configure it to be a skill in a Copilot agent.

> [!NOTE]
>
> - Currently, only Excel, PowerPoint, and Word add-ins can be configured as a skill in Copilot. We're working to support Outlook.
> - Copilot agents are not currently supported in Office on Mac.
> - An add-in must use the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md) to be configured as a skill in Copilot.
> - A [content add-in](content-add-ins.md) cannot be a skill in Copilot.

## Major tasks

There are two major tasks to configuring an add-in as a Copilot skill, and they are analogous to the two tasks for configuring [function commands](add-in-commands.md#types-of-add-in-commands) for an add-in.

- Create JavaScript functions that implement the add-in's actions.
- Use JSON to specify for Office and the JavaScript runtimes the names of these functions.

## JSON configuration

Configuration of an add-in to be a Copilot skill requires three JSON-formatted files that are described in the following subsections.

### Unified manifest for Microsoft 365

There are two parts of the manifest that you configure. First, create an action object that identifies the JavaScript function that is invoked by the action. The following is an example (with some extraneous markup omitted). NOte the following about this code.

- The "page" property specifies the URL of the web page that contains an embedded script tag that specifies the URL of the JavaScript file where the function is defined. That same file contains an invocation of the [Office.actions.associate](/javascript/api/office/office.actions?view=common-js-preview#office-office-actions-associate-member(1)) method to map the function to an action ID.
- The "actions.id" property in the manifest is the same action ID that is passed to the call of `associate`.
- The "actions.type" property is set to "executeDataFunction", which is the type that can accept parameters and can be invoked by Copilot.

```json
"extensions": [

    ...

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
                    "id": "fillcolor",
                    "type": "executeDataFunction",
                }
            ]
        }
    ]
]
```

Second, create a declarative agent object that identifies the file containing the detailed configuration of the agent. The following is an example.

```json
"copilotAgents": {
    "declarativeAgents": [
        {
        "id": "ContosoCopilotAgent",
        "file": "declarativeCopilot.json"
        }
    ]
}
```

The reference documentation for the manifest JSON is at [Public developer preview app manifest](/microsoft-365/extensibility/schema/?view=m365-app-prev).

### Declarative agent configuration

The agent configuration file includes instructions for the agent and specifies one or more API plugin configuration files that will contain the detailed configuration of the agent's actions. The following is an example. Note the following about this JSON.

- The conversation starter appears in the chat canvas of Copilot.
- The "actions.id" property in this file is the collective ID of all the functions in the file specified in "actions.file". It doesn't have to match the "actions.id" in the manifest.

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

The reference documentation for declarative agents is at [Declarative agent schema 1.3 for Microsoft 365 Copilot](/microsoft-365-copilot/extensibility/declarative-agent-manifest-1.3).

### Copilot API plugin configuration

The API plugin configuration file specifies the "functions" of the plugin in the sense of agent actions, not JavaScript functions, including the instructions for the action. It also configures the JavaScript runtime for Copilot. The following is an example. About this JSON, note the following:

- The "functions.name" must match the "extensions.runtimes.actions.id" property in the add-in manifest.
- The "reasoning.description" and "reasoning.instructions" refer to a JavaScript function, not a REST API.
- The "runtimes.run_for_functions" array must include either the same string as "functions.name" or a wildcard string that matches it.
- The "runtimes.spec.local_endpoint" property is new and is not yet in the main reference documentation for the API plugins schema. See below for more about it. In this case, it specifies that the JavaScript function that is associated with the "fillcolor" string is available in an Office add-in, rather than in some REST endpoint.

```json
{
    "schema_version": "v2.2",
    "name_for_human": "Excel Add-in + Agent",
    "description_for_human": "Add-in Actions in Agents",
    "namespace": "addin_function",
    "functions": [
        {
            "name": "fillcolor",
            "description": "fillcolor changes a single cell location to a specific color.",
            "parameters": {
                "type": "object",
                "properties": {
                    "Cell": {
                        "type": "string",
                        "description": "A cell location in the format of A1, B2, etc.",
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
                    "instructions": "The user will ask for a color that isn't in the hex format needed in most cases, make sure to convert to the closest approximation in the right format."
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

The reference documentation for API plugins is at [API plugin manifest schema 2.2 for Microsoft 365 Copilot](/microsoft-365-copilot/extensibility/api-plugin-manifest-2.2). The following is the documentation for the "runtimes.spec.local_endpoint" property.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `local_endpoint` | String | Optional. The ID of a set of available JavaScript functions. This property is roughly analogous to the ["runtimes.spec.url"](/microsoft-365-copilot/extensibility/api-plugin-manifest-2.2#openapi-specification-object) property, but for local functions on the client, not REST APIs.|

## Create the JavaScript functions

The JavaScript functions that will be invoked by the Copilot agent are created exactly as [function commands](../develop/create-addin-commands-unified-manifest.md#add-a-function-command) are created. The following is an example. Note the following about this code.

- Unlike a function command, a function associated with a Copilot action can take parameters.
- The first parameter of the `associate` method must match both the "extensions.runtimes.actions.id" property in the add-in manifest and the "functions.name" property in the API plugins JSON.

```javascript
async function fillColor(cell, color) {
    await Excel.run(async (context) => {
        context.workbook.worksheets.getActiveWorksheet().getRange(cell).format.fill.color = color;
        await context.sync();
    })
}

Office.onReady((info) => {
    Office.actions.associate("fillcolor", async (message) => {
        const {cell, color} = JSON.parse(message);
        await fillColor(cell, color);
        return "Cell color changed.";
    });
});
```

After your functions are created, create a UI-less HTML file that contains a `<script>` tag that loads the JavaScript file with the functions. The URL of this HTML file must match the value of the "extensions.runtimes.code.page" property in the manifest. See [Unified manifest for Microsoft 365](#unified-manifest-for-microsoft-365) earlier in this article.
