---
title: Configure your Office Add-in to use a shared runtime
description: Configure your Office Add-in to use a shared runtime to support additional ribbon, task pane, and custom function features.
ms.topic: how-to
ms.date: 07/30/2026
ms.localizationpriority: high
---

# Configure your Office Add-in to use a shared runtime

On desktop platforms, by default, your add-in runs code for ribbon buttons, custom functions, and the task pane in separate runtime environments. This creates limitations, such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.

However, you can configure your Office Add-in to share code in the same runtime (referred to as a shared runtime). This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.

Configuring a shared runtime enables the following scenarios.

- Your Office Add-in can use additional UI features.
  - [Change the availability of add-in commands](../design/disable-add-in-commands.md)
  - [Run code in your Office Add-in when the document opens](run-code-on-document-open.md)
  - [Show or hide the task pane of your Office Add-in](show-hide-add-in.md)
  - [Add custom keyboard shortcuts to your Office Add-ins](../design/keyboard-shortcuts.md) (supported in Excel, PowerPoint, and Word add-ins)
- The following are available for Excel add-ins only.
  - [Create custom contextual tabs in Office Add-ins](../design/contextual-tabs.md)
  - Custom functions have full CORS support.
  - Custom functions can call Office.js APIs to read spreadsheet document data.

> [!NOTE]
> The runtime in which the [Office Dialog](dialog-api-in-office-add-ins.md) runs can't be shared, but that is not a significant limitation. Consider the following.
>
> - Use the [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) and [messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) functions to instantly communicate between a dialog runtime and a shared runtime. Doing so creates an experience for the user that is the same as it would be if the dialog is running in the same runtime.
> - The function that opens the Office dialog can be passed a parameter that causes the dialog to use the same runtime as a parent task pane, but only when the add-in is running in Office on the web. 
>
> For more information, see [Use the Office dialog API in Office Add-ins](dialog-api-in-office-add-ins.md).

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

This article steps through the process of configuring an add-in to use a shared runtime.

> [!TIP]
> If the add-in was created with the option for an Excel custom function in either [Microsoft 365 Agents Toolkit](../develop/teams-toolkit-overview.md) or the [Yeoman generator for Office Add-ins](yeoman-generator-overview.md), then it is already configured to use a shared runtime.

## Create the add-in project

To learn how to *convert* an add-in to use a shared runtime, start by creating an add-in project that isn't already configured to used a shared runtime to use as a continuing example. Use the Agents Toolkit to create a *task pane* project, not a custom function project. Instructions are in [Create Office Add-in projects with Microsoft 365 Agents Toolkit](agents-toolkit-overview.md).

> [!NOTE]
> This article uses filenames that are in the continuing example, and are common in Office add-ins; **taskpane**, **commands**, and **functions**. If you are configuring an existing add-in that uses different filenames, treat these filenames as placeholders. 

## Configure the manifest

Follow these steps to configure a project to use a shared runtime. The continuing example project uses the unified manifest.  

> [!IMPORTANT]
> If you are converting an existing add-in that uses the add-in only manifest, then open that tab. But we recommend that you first convert your add-in to use the unified manifest, and then configure it to have a shared runtime.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

1. Open your add-in project in Visual Studio Code.
1. Open the **\appPackage\manifest.json** file.
1. Replace the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array?view=m365-app-prev&preserve-view=true) array with the following JSON. Note the following about this markup.
    - The [SharedRuntime 1.1 requirement set](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets#sharedruntime-api-11) is specified in the [`"requirements.capabilities"`](/microsoft-365/extensibility/schema/requirements-extension-element-capabilities) object. This configures your add-in to run in a shared runtime on supported clients. For a list of clients that support the SharedRuntime 1.1 requirement set, see [Shared runtime requirement sets](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets).
    - The `"id"` of the runtime is set to the descriptive name `"SharedRuntime"`.
    - The `"lifetime"` property is set to `"long"`. This is the setting that makes the add-in use a shared runtime. The default value of `"lifetime"` is `"short"`.

      > [!NOTE]
      > If you are configuring an existing add-in to use a shared runtime, and it has more than one object in the `"runtimes"` array, only one runtime object may have its `"lifetime"` property is set to `"long"`.  

    ```json
    "runtimes": [
      {
        "requirements": {
            "capabilities": [
                { 
                    "name": "AddinCommands", 
                    "minVersion": "1.1" 
                },
                {
                    "name": "SharedRuntime",
                    "minVersion": "1.1"
                }
            ]
        },
        "id": "SharedRuntime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/taskpane.html"
        },
        "lifetime": "long",
        "actions": [
          {
            "id": "TaskPaneRuntimeShow",
            "type": "openPage"
          },
          {
            "id": "action",
            "type": "executeFunction"
          }
        ]
      }
    ]
    ```

1. Save your changes.

# [Add-in only manifest](#tab/xmlmanifest)

1. Open your add-in project in Visual Studio Code.
1. Open the **manifest.xml** file.
1. Update the requirements section to include the [shared runtime](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets) as follows.

    ```xml
    <Hosts>
      <Host ...>
    </Hosts>
    <Requirements>
        -- possibly other requirements here --
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    ```

1. Find the `<VersionOverrides>` section and add the following `<Runtimes>` section. Note the following about this markup.

    - The `lifetime` attribute is set to **long**. This is the setting that makes the add-in use a shared runtime. The default value of `lifetime` is **short**.
    - The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the `<bt:Urls>` section near the bottom of the **manifest.xml** file.

        > [!IMPORTANT]
        > The shared runtime won't load if the `resid` of the runtime isn't consistent in the manifest. If you change the value to something other than **Taskpane.Url**, be sure to also change the value in all locations shown in the following steps in this article.

    - The `<Runtimes>` section must be entered after the `<Host>` element in the exact order shown in the following XML.

    ```xml
    <VersionOverrides ...>
      <Hosts>
        <Host ...>
          <Runtimes>
            <Runtime resid="Taskpane.Url" lifetime="long" />
          </Runtimes>
        ...
        </Host>
    ```

1. If the add-in is an Excel add-in with custom functions, find the `<Page>` element. Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Find the `<FunctionFile>` tag and change the `resid` from **Commands.Url** to  **Taskpane.Url**. Note that if you don't have function commands, you won't have a `<FunctionFile>` entry, and can skip this step.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Save your changes.

---

## Configure the webpack.config.js file

In the continuing example, and probably in an existing add-in, the **webpack.config.js** builds multiple runtime loaders. You need to modify it to load only the shared runtime via the **taskpane.html** file.

1. Open the **webpack.config.js** file.
1. If your **webpack.config.js** file has the following **commands.html** plugin code, remove it.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. If you're configuring an existing custom functions add-in to use a shared runtime, then the **webpack.config.js** file has the following **functions.html** plugin code, remove it.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. If your project used either the **functions** or **commands** chunks, add them to the chunks list for the task pane as shown in the following code.

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Save your changes and rebuild the project.

   ```command&nbsp;line
   npm run build
   ```

> [!NOTE]
> If your project has a **functions.html** file or **commands.html** file, they can be removed. The **taskpane.html** loads the **functions.js** and **commands.js** code into the shared runtime via the webpack updates you just made.

## Test your Office Add-in changes

Confirm that you're using the shared runtime correctly with the following steps.

1. Open the **taskpane.js** file.
1. Comment out the existing code in the file and then add the following code. This code displays a count of how many times the task pane has been opened. Adding the `onVisibilityModeChanged` event is only supported in a shared runtime.

    ```javascript
    /*global document, Office*/

    let _count = 0;

    Office.onReady(() => {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";

      updateCount(); // Update count on first open.
      Office.addin.onVisibilityModeChanged((args) => {
        if (args.visibilityMode === Office.VisibilityMode.taskpane) {
          updateCount(); // Update count on subsequent opens.
        }
      });
    });

    function updateCount() {
      _count++;
      document.getElementById("run").textContent = "Task pane opened " + _count + " times.";
    }
    ```

1. Save your changes and run the project.

   ```command&nbsp;line
   npm start
   ```

Each time you open the task pane, the count of how many times it has been opened is incremented. The value of **_count** isn't be lost because the shared runtime keeps your code running even when the task pane is closed.

When you're finished testing, follow best practices to stop the dev server and uninstall the add-in as described at [Use your tool's uninstall facility](../testing/uninstall-add-in.md#use-your-tools-uninstall-facility). Then restore the original code of **taskpane.js**. 

## Good practice: avoid multiple task panes

A shared runtime only supports one task pane, although that task pane can have more than one page. Implementing this practice depends on the type of manifest.

- **Unified manifest for Microsoft 365**: In the runtime object that is configured with a `"long"` lifetime, none of the objects in the `"actions"` array should have a `"view"` property. 
- **Add-in only manifest**: In the ancestor `<Host>` element that has a descendant `<Runtime>` element that is set to a `long` lifetime, none of the descendant `<Action>` elements should have a `<TaskpaneID>` child element.

## See also

- [Call Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md)
- [Add custom keyboard shortcuts to your Office Add-ins](../design/keyboard-shortcuts.md)
- [Create custom contextual tabs in Office Add-ins](../design/contextual-tabs.md)
- [Change the availability of add-in commands](../design/disable-add-in-commands.md)
- [Run code in your Office Add-in when the document opens](run-code-on-document-open.md)
- [Show or hide the task pane of your Office Add-in](show-hide-add-in.md)
- [Tutorial: Share data and events between Excel custom functions and the task pane](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
