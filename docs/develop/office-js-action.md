---
title: Add an Office JavaScript Library action to Copilot
description: Learn how to enable Copilot to invoke an Office JavaScript Library action.
ms.date: 10/19/2026
ms.topic: how-to
ms.service: microsoft-365
ms.localizationpriority: medium
---

# Add an Office JavaScript Library action to Copilot

In this tutorial, you'll create an action that enables Copilot to do something in an open Excel workbook that Copilot doesn't know how to do natively. 

> [!NOTE]
> This article assumes that you're familiar with the overview [Combine Copilot with Office.js actions](../design/office-js-action-overview.md). We also recommend that you complete the quick start [Build your first Office.js action as a Copilot skill](../quickstarts/office-js-action-quickstart.md).

> [!IMPORTANT]
> This feature requires the [unified manifest for Microsoft 365](unified-manifest-overview.md). If your project uses the Office Add-in only manifest, you must first [convert it to use the unified manifest](convert-xml-to-json-manifest.md) before you can enable Copilot to invoke an Office.js action.

## Major tasks

The following are the main tasks for enabling Copilot to invoke an Office JavaScript Library (Office.js) action in your project. Details are in the subsections.

- [Create the functions for the actions](#create-the-functions-for-the-actions) that implement the actions.
- [Update the manifest](#update-the-manifest)
- [Create the action file](#create-the-action-file)
- [Create the app package](#create-the-app-package)
- [Test the action](#test-the-action)
- [Make changes in the project](#make-changes-in-the-project)

### Create the functions for the actions

If your project includes one or more [function commands](../design/add-in-commands.md#types-of-add-in-commands), then your project already has a JavaScript or TypeScript file that defines the functions for these commands (usually called **commands.js** or **commands.ts**) and a UI-less HTML file (usually called **commands.html**) that has a `<script>` tag to load the function file. We recommend that you use this same function file to define the functions for your Office.js actions. Skip to the section [Update the function file](#update-the-function-file).

#### Create the source files

If your project doesn't already have such files, then create them with the following steps. The folder and file structure and names in these steps aren't mandatory, but we recommend them to minimize incompatibilities with development tooling and configuration for bundlers and transpilers where applicable.

1. Ensure that there's a folder **\src** off the root of the project, and that it has a child folder named **\commands**.
1. Create a **commands.js** (or .ts) file in the **\commands** folder. You add content to it in a later step.
1. Create a **commands.html** file in the **\commands** folder with the following content. Note the following about this markup.

   - The `<body>` element is empty because the file has no UI. Its only purpose is to load JavaScript files.
   - The Office JavaScript Library and the **commands.js** file that you created in the preceding step is explicitly loaded.

       > [!NOTE]
       > It's common in Office.js development to use tools like [webpack](https://webpack.js.org/) and its plug-ins to automatically inject `<script>` tags into HTML files at build time. If you use such a tool, you shouldn't include any `<script>` tags in your source file that are going to be inserted by the tool.

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

Define the functions for your Office.js actions with the following steps. These assume the most common folder and file structure and names.

Open the function file (usually, the **\src\commands\commands.js** (or .ts)) and add the functions that implement the actions. Keep the following points in mind as you work.

- Like all code that calls APIs in the Office JavaScript Library, the file must [initialize the library](initialize-add-in.md), usually by calling `Office.onReady`.
- Functions that are invoked by Copilot take a `message` parameter that specifies the values that Office uses when it calls functions from the Office JavaScript Library. This object can be parsed with the `JSON.parse` method. The values of the object are specified by users in natural language.
- Unlike the functions that are invoked with function commands, these functions don't have an `Office.AddinCommands.Event` parameter, and they don't call the `Office.AddinCommands.Event.completed` method. Instead, functions that are invoked by Copilot return a result object from the JavaScript runtime back to the Copilot runtime.
- Functions invoked from Copilot have less time to complete than functions invoked from a function command. The latter have five minutes to complete or Office shuts down the JavaScript runtime. But Copilot-invoked functions have only two minutes to return a result or the runtime is shut down.
- For each function, there must be a call of [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) to tell Office which function in the file should be run when the action is invoked. The `associate` function maps the function name (or an anonymous function) to an action ID that you configure in the manifest in a later step. If you define multiple functions in the same file, your code must call `associate` for each one.

The following is an example.

```javascript
Office.onReady(function() {
    // Add any initialization code here.
});

async function applyContosoFinancialStyleFromUserData(message) {
    // message is { rangeAddress: "A1:C4" }, for some range address
    const {rangeAddress} = JSON.parse(message);
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
    return `Applied the Contoso financial reporting style to ${rangeAddress}.`;
}

Office.actions.associate("applyContosoFinancialStyle", applyContosoFinancialStyleFromUserData);
```

### Update the manifest

There are two parts to configuring the manifest, as described in the following subsections.

#### Use the preview manifest schema

Ensure that the manifest references the preview version of the manifest schema with the following steps.

1. At the top of the manifest, ensure that the [`"$schema"`](/microsoft-365/extensibility/schema/root#schema-4) property is the following:

   ```json
   "$schema": "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
   ```

1. Ensure that the [`"manifestVersion"`](/microsoft-365/extensibility/schema/root#manifestversion) property is set to "devPreview".

#### Configure the runtime

1. In the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array) array, ensure that there's a runtime object that is oriented to running UI-less JavaScript functions. It's critical that the object have a [`"code.page"`](/microsoft-365/extensibility/schema/extension-runtime-code#page) property that specifies the absolute URL to the UI-less HTML file that you created (or edited) earlier in [Update the function file](#update-the-function-file).
1. In the [`"actions"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item) array of that same runtime object, add an object with the following properties.
   - `"id"`: an exact match for the first parameter of the call to `Office.actions.associate` in a function that you created in [Update the function file](#update-the-function-file).
   - `"type"`: set to "executeDataFunction".
   - `"actionFile"`: the path, relative to the manifest, to the action file you create in [Create the action file](#create-the-action-file).
1. Repeat the preceding step for each function you created to implement an action.

The runtime object should look similar to the following. There may be other properties of the runtime object, other runtime objects in the `"runtimes"` array, and other action objects in the `"actions"` array.

```json
"runtimes": [
    {
        "id": "CommandsRuntime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/commands.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "applyContosoFinancialStyle",
                "type": "executeDataFunction",
                "actionFile": "./office-actions/ApplyContosoFinancialStyleAction.json"
            }
        ]
    }
]
```

### Create the action file

1. Create a folder in the same location as your manifest to hold your action files; for example, **office-actions**.
1. Create a file in that folder and give it the name used in the `"actions.actionFile"` property, such as **ApplyContosoFinancialStyleAction.json**.
1. Paste the following content into the file.

   ```json
   {
       "$schema": "https://developer.microsoft.com/json-schemas/office-js/1.0/action.schema.json",
       "id": "applyContosoFinancialStyle",
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

1. *Replace the property values with new values that are appropriate for your project.* For more information about these properties, see [Action file object](action-file-schema-1.0.md#action-file-object).

   As you work, keep the following points in mind.

   - The `"id"` property must be an exact match for both of the following:

      - An `"extensions.runtimes.actions.id"` property in the manifest (for an action of type "executeDataFunction").
      - The first parameter of the call to `Office.actions.associate` in a function that you created in [Update the function file](#update-the-function-file).

   - Add additonal parameter objects to the `"parameters"` array as needed. Copilot passes to the JavaScript function an object with each parameter name and value; for example, `{ rangeAddress: "A1:C4", color: "blue" }`
   - Add a `"states.reasoning"` property if appropriate.

### Create the app package

> [!IMPORTANT]
> To test the action, as described later in this article in [Test the action](#test-the-action), the domain segment of any absolute URLs in the manifest must be a localhost domain; for example, `localhost:3000`. You can change these segments to a production domain later.

Using any zip utility, create a zip file that contains the following files.

- The manifest.
- The icon files referenced in the manifest's [`"icons.color"`](/microsoft-365/extensibility/schema/root-icons#color) and [`"icons.outline"`](/microsoft-365/extensibility/schema/root-icons#outline).
- The action file (or files) you created in [Create the action file](#create-the-action-file).
- Any files referenced in the [`"localizationInfo.additionalLanguages"`](/microsoft-365/extensibility/schema/root-localization-info#additionallanguages) property.

> [!IMPORTANT]
> Most of these files have URLs in the manifest that are relative to the location of the manifest, so the folder structure of the zip file must maintain this structure. For example, if the `"icons.color"` value is "/assets/icon-32.png", then there must be an **/assets** folder in the zip file with the **icon-32.png** file in it. If the `"actionFile"` value is "/office-actions/myAction.json", then there must be an **/office-actions** folder in the zip file with the **myAction.json** file in it.

> [!TIP]
> To maximize compatibility with Microsoft development tools, we recommend that you create a subfolder named **build** in the **appPackage** folder and create the zip file in it.

### Test the action

The three major steps to test the action &mdash; sideloading, starting a server, and running the action &mdash; are described in the following subsections.

#### Sideload the project

Sideloading is done through Teams even when there's no Teams feature in the project. The steps are as follows.

1. Close all Office applications, and then clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select the gear icon in the upper right.
1. Select **Upload an app** in the **Apps** dialog, and then in the dialog that opens, select **Upload a custom app**.
1. In the **Open** dialog, navigate to, and select, the package zip file.
1. Select **Add** in the dialog that opens.
1. When you're prompted that the app was added, *don't* open it in Teams. Instead, close Teams.

#### Start the server

Your task is to start a local web server that hosts your project's HTML and JavaScript files. How you do this depends on several factors including the folder structure of your project, the tools you use, such as a bundler, task manager, server application, and how you have configured those tools. Since you're adding an action to an existing project, you already know how to do this. The following instruction applies only to projects that meet the following conditions.

- There's a **webpack.config.js** file in the root of the project that is similar to the ones in Office Add-in projects that are created with the [Microsoft 365 Agent Toolkit](teams-toolkit-overview.md).
- There's a **package.json** file in the root of the project similar to the ones created by Agent Toolkit and the file has a "scripts" section with the following script in it.

   ```json
   "dev-server": "webpack serve --mode development"
   ```

In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run dev-server` to start the server on localhost.

#### Run the action

1. Open the Excel. Wait until the project has loaded. This may take as much as three minutes. Depending on your version of Office, ribbon buttons and other artifacts may appear automatically. If they don't, you need to manually activate the project by selecting the **Add-ins** button on the **Home** tab, and then selecting your app from the flyout that opens.
1. Select the dynamic action button for Copilot Chat on the surface of the open document.

    :::image type="content" source="../images/copilot-dynamic-action-button.png" alt-text="The dynamic action button for Copilot: a round button with the Copilot symbol. There is no text on the button.":::

1. Try entering prompts in the conversation box that your action should be able to handle. Remember that Copilot can natively make almost any kind of read or write operation to the open document, so be sure that your prompt mentions your custom action. For example, "Apply Contoso's required formatting to the range A1:C4.", not "Format the range A1:C4." 

### Make changes in the project

Live reloading and hot reloading aren't supported in the preview period. To make changes, first shut down the server and uninstall the project with these steps.

1. Close the Office application.
1. If the web server is running in the Visual Studio Code **TERMINAL**, give the terminal focus and press <kbd>Ctrl</kbd>+<kbd>C</kbd>. Choose "Y" in response to the prompt to end the process. Then go to the next step. If the web server is running in a separate window, skip this step and go to the next step.
1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run stop`.
1. Clear the Office cache following the instructions at [Manually clear the cache](../testing/clear-cache.md#manually-clear-the-cache).
1. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
1. Find your project in the list of apps.
1. Select the arrow head to the left of the name to expand its row.
1. Select the trash can icon near the right end of the row, and then select **Remove** in the prompt.

Make your changes and then repeat the steps in [Test the action](#test-the-action).

## Troubleshooting

See [Troubleshooting Office.js actions in Copilot](../design/office-js-action-overview.md#troubleshooting-officejs-actions-in-copilot).
