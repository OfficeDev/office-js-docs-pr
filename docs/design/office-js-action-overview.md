---
title: Combine Copilot with Office JavaScript Library actions (preview)
description: Get an overview of why and how to enable Copilot to invoke Office.js actions.
ms.date: 10/19/2026
ms.topic: overview
ms.localizationpriority: medium
---

# Combine Copilot with Office JavaScript Library actions (preview)

Office JavaScript Library (Office.js) *actions* are JavaScript functions that are invoked from Copilot in Office and call [Office.js APIs](../develop/understand-the-javascript-api-for-office.md) to read and write to an open Office document. Enabling Copilot to invoke Office.js actions provides two benefits:

- Copilot becomes a natural language interface for your JavaScript functions.
- Copilot can pass parameters to the JavaScript it invokes, which isn't possible when a [function command](add-in-commands.md#types-of-add-in-commands) is invoked from a button or menu item.

> [!NOTE]
> Office.js actions are in preview and only supported in Excel. Don't use them in a production app for Microsoft 365.

## Scenarios

The following are some selected ways in which enabling Copilot to invoke Office.js actions enhances the value of an app to users.

- **Learning how to use the app's functionality**: When a user needs to perform multiple steps or tasks to achieve a goal, the chat interface of Copilot can ease the process of getting started. For example, consider a legal firm that needs to have a list of questions that must be answered about each lease that it prepares. Creating this list of questions can be time-consuming and labor-intensive. But Copilot, using an Office.js action, can be prompted to produce a first draft list of questions and insert them into a Word document.

- **Content analysis**: Copilot can be used to analyze the content of a document or spreadsheet and take action depending on what it finds. The following are examples.

   - Copilot analyzes a Request for Proposal and then fetches the answers to questions in the RFP from a backend system. The user simply prompts Copilot to "Fill in the answers you know to the questions."
   - Copilot analyzes a document, or a table in a spreadsheet, for content that implies certain actions must be taken, either in the document itself or elsewhere in the customer's business systems. The user might say "Review the document for any items I missed on the audit list."

- **Trusted insertion of data**: If you prompt a typical AI engine with a question, it combines information it finds and compose an answer; a process that can introduce inaccuracies. But Copilot, invoking an Office.js action, can insert data *unchanged* from a trusted source. Some examples:

   - Consider an Office.js action that enables the insertion of legal research into Word where it can then be edited. A user prompts Copilot: "In what circumstances can a lease of residential space in Indiana be broken unilaterally by the lessor?" The action then fetches content, unchanged, from precedents and statutes.
   - Consider an Office.js action that manages the inventory of digital assets. In Copilot chat, a user prompts: "Insert a table of our color photos with the name of each, the number of times it was downloaded, and its size in megabytes, sorted in order from most downloaded." The action then fetches this data, unchanged, from the system of record and inserts the table into an Excel spreadsheet.

## The relation of Copilot to Office.js

Copilot is a natural language interface for the JavaScript functions in your app for Microsoft 365. A app can be configured to expose *only* Office.js actions to Copilot. It doesn't have to include an Office Add-in with components such as task pane, custom ribbon buttons, custom menus, event handlers, or custom functions; but it can have any of these in addition to its Office.js actions. The best approach depends on the user scenarios that the app should enable.

- If the app should provide simple, fast actions in Office that don't need parameters passed to them, include an add-in with custom ribbon buttons or menu items, called [function commands](add-in-commands.md#types-of-add-in-commands).
- If the app needs a dashboard experience, needs the user to configure settings, needs to display metadata about the content of the Office document, or needs a page-like experience for any other reason, include an add-in with a task pane.
- If the app needs to start up in Office in response to an event, such as a document opening, include an add-in with a handler for the event.
- If the app needs to provide Excel users with custom formulas they can insert in worksheet cells, include an Excel add-in with custom functions.
- If the app needs to provide complex actions that require parameters passed at runtime or needs a natural language interface, include Office.js actions for Copilot.

> [!NOTE]
>
> - Currently, only Excel can expose Office.js actions to Copilot. We're working to support other Word, PowerPoint, and Outook.
> - A app must use the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md) to expose Office.js actions to Copilot.

## Major tasks

There are two major tasks to enable Copilot to invoke your app's Office.js actions, and they're analogous to the two tasks for configuring [function commands](add-in-commands.md#types-of-add-in-commands) for a app.

- Create JavaScript functions that implement the actions.
- Use JSON to specify the names and behavior of these functions.

## Configuration

Configuration is in one or more JSON-formatted action files that are included in the app package zip file. References to these files are included in the app manifest.

### Manifest

In the manifest, create an action object that identifies the JavaScript function that's invoked by the action, and points to the action file that describes it. The following is an example. Note the following about this code.

- The "page" property specifies the URL of the web page that contains an embedded `<script>` tag that, in turn, specifies the URL of the JavaScript file where the function is defined. That JavaScript file calls the [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) method to map the function to an action ID.
- The [`"actions.id"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#id) property in the manifest is the same action ID that's passed to the call of `associate`.
- The [`"actions.type"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#type) property is set to `"executeDataFunction"`, which is the type that can accept parameters and can be invoked by Copilot.
- The [`"actions.actionFile"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item#actionFile) property is the path, relative to the manifest, to the action file that describes this action. The action file must conform to the [action file schema](../develop/action-file-schema-1.0.md).

   > [!NOTE]
   > During the preview period for Office.js actions, you must use the [preview version of the unified manifest for Microsoft 365](/microsoft-365/extensibility/schema/?view=m365-app-prev&preserve-view=true).

```json
"extensions": [

    ...

    "runtimes": [
        {
            "id": "ActionsRuntime",
            "type": "general",
            "code": {
                "page": "https://localhost:3000/commands.html",
                "script": "https://localhost:3000/commands.js"
            },
            "lifetime": "short",
            "actions": [
                {
                    "id": "ApplyContosoFinancialStyle",
                    "type": "executeDataFunction",
                    "actionFile": "./office-actions/ApplyContosoFinancialStyleAction.json"
                }
            ]
        }
    ]
]
```

### Action file

An action file describes a single Office.js action: its ID, a description of what it does and when to use it, its parameters, its return value, and optional per-state instructions for the Copilot orchestrator.  The following is an example. Note the following about this JSON.

- The `"id"` property must be an exact match for the `"actions.id"` property of the corresponding action object in the manifest, and for the first parameter of the call to `Office.actions.associate` in the JavaScript function that implements the action.
- Copilot passes all the parameters in the `"parameters"` array as a JSON object, even when there is only one, as in this example. So, the JavaScript function that is called by the action receives an object like `{ "rangeAddress": "A1:A12" }` as the parameter.
- The `"states.responding"` property only provides *guidance* to Copilot about how to respond. It doesn't put any limits or structural requirements on the response.

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/1.0/action.schema.json",
    "id": "ApplyContosoFinancialStyle",
    "actionDescription": "Applies the Contoso financial reporting style to an Excel range, including Contoso's required header colors, fonts, alignment, borders, and number formatting.",
    "parameters": [
        {
            "name": "rangeAddress",
            "parameterDataType": "string",
            "parameterDescription": "The address of the Excel range to format, such as A1:F12.",
            "isRequired": true
        }
    ],
    "returnedValue": {
        "dataType": "string",
        "returnedValueDescription": "A string indicating the result of the action."
    },
    "states": {
        "responding": {
            "stateDescription": "If there is no error, tell the user which range was formatted with the Contoso financial reporting style."
        }
    }
}
```

The reference documentation for action files is at [Action file schema 1.0 for Office.js actions](../develop/action-file-schema-1.0.md).

## Create the JavaScript functions

The JavaScript functions that are invoked by Copilot are defined in a JavaScript file just as [function commands](../develop/create-addin-commands-unified-manifest.md#add-a-function-command) are created. 

> [!TIP]
> If you have both Office.js actions and an add-in with function commands in the same app, you can use the same JavaScript file for both the actions and the function commands.

The following is an example. Note the following about this code.

- Unlike a function command, a function associated with an Office.js action can takes parameters, in the form of a single JSON `message` object. The object can be parsed and its members passed to helper functions. In this example, the `message` object is something like `{ "rangeAddress": "A1:A12" }`.
- The first parameter of the `associate` method must match both the `"extensions.runtimes.actions.id"` property in the manifest and the `"id"` property in the action file.

```javascript
Office.onReady(() => {
    Office.actions.associate("ApplyContosoFinancialStyle", async (message) => {
        const {rangeAddress} = JSON.parse(message);
        await applyContosoFinancialStyle(rangeAddress);
        return `Applied the Contoso financial reporting style to ${rangeAddress}.`;
    });
});

async function applyContosoFinancialStyle(rangeAddress) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(rangeAddress);
        const header = range.getRow(0);

        range.format.font.name = "Aptos";
        header.format.fill.color = "#0F6CBD";
        header.format.font.color = "#FFFFFF";
        header.format.font.bold = true;
        header.format.horizontalAlignment = "Center";
        range.format.autofitColumns();

        await context.sync();
    });
}
```

After your functions are created, create a UI-less HTML file that contains a `<script>` tag that loads the JavaScript file with the functions. The URL of this HTML file must match the value of the [`"extensions.runtimes.code.page"`](/microsoft-365/extensibility/schema/extension-runtime-code#page) property in the manifest. See [Manifest](#manifest) earlier in this article.

## Troubleshooting Office.js actions in Copilot

The following are some possible problems and suggested solutions.

- The action fails with a message indicating that the action wasn't found. The following are some possible causes.

   - There's no `"extensions.runtimes.actions"` object in the manifest with an `"id"` that exactly matches the action file's `"id"` property.
   - There is a matching `"actions.id"` in the manifest, but the sibling `"actions.type"` value for the same action object isn't `"executeDataFunction"`.
   - The `"actions.actionFile"` property doesn't point to a valid action file, or doesn't have the correct path to the file relative to the manifest in the app package.

- The action fails with a message indicating the action handler registration wasn't found. The following is a possible cause.

   - The app's JavaScript doesn't have a call of `Office.actions.associate` with the first parameter *exactly matching* the `"id"` property value in the action file.

## Next steps

- [Build your first Office.js action](../quickstarts/office-js-action-quickstart.md)
- [Add an Office.js action to Copilot](../develop/office-js-action.md)
