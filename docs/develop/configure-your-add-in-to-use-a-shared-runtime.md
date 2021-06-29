---
ms.date: 06/14/2021
title: "Configure your Office Add-in to use a shared JavaScript runtime"
ms.prod: non-product-specific
description: 'Configure your Office Add-in to use a shared JavaScript runtime to support additional ribbon, task pane, and custom function features.'
localization_priority: Priority
---

# Configure your Office Add-in to use a shared JavaScript runtime

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

You can configure your Office Add-in to run all of its code in a single shared JavaScript runtime (also known as a shared runtime). This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in. It also enables additional features such as running code when the document opens, or enabling or disabling ribbon buttons. To configure your add-in to use a shared JavaScript runtime, follow the instructions in this article.

## Create the add-in project

If you are starting a new project, follow these steps to use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create an Excel or PowerPoint add-in project.

Do one of the following:

- To generate an Excel add-in with custom functions, run the command `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.

    or

- To generate a PowerPoint add-in, run the command `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`.

The generator will create the project and install supporting Node components.

> [!NOTE]
> You can also use the steps in this article to update an existing Visual Studio project to use the shared runtime. However, you may need to update the XML schemas for the manifest. For more information, see [Troubleshoot development errors with Office Add-ins](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

## Configure the manifest

Follow these steps for a new or existing project to configure it to use a shared runtime. These steps assume you have generated your project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).

1. Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.
1. Open the **manifest.xml** file.
1. If you generated an Excel add-in, update the requirements section to use the [shared runtime](../reference/requirement-sets/shared-runtime-requirement-sets.md) instead of the custom function runtime. The XML should appear as follows.

    ```xml
    <Hosts>
      <Host Name="Workbook"/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. Find the `<VersionOverrides>` section and add the following `<Runtimes>` section just inside the `<Host ...>` tag. The lifetime needs to be **long** so that your add-in code can run even when the task pane is closed. The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the ` <bt:Urls>` section near the bottom of the **manifest.xml** file.

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
       <Runtimes>
         <Runtime resid="Taskpane.Url" lifetime="long" />
       </Runtimes>
       ...
   ```

1. If you generated an Excel add-in with custom functions, find the `<Page>` element. Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Find the `<FunctionFile ...>` tag and change the `resid` from **Commands.Url** to  **Taskpane.Url**. Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Save the **manifest.xml** file.

## Configure the webpack.config.js file

The **webpack.config.js** will build multiple runtime loaders. You need to modify it to load only the shared JavaScript runtime via the **taskpane.html** file.

1. Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.
1. Open the **webpack.config.js** file.
1. If your **webpack.config.js** file has the following **functions.html** plugin code, remove it.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. If your **webpack.config.js** file has the following **commands.html** plugin code, remove it.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. If your project used either the **functions** or **commands** chunks, add them to the chunks list as shown next (the following code is for if your project used both chunks).

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Save your changes and rebuild the project.

   ```command line
   npm run build
   ```

> [!NOTE]
> If your project has a **functions.html** file or **commands.html** file, they can be removed. The **taskpane.html** will load the **functions.js** and **commands.js** code into the shared JavaScript runtime via the webpack updates you just made.

## Test your Office Add-in changes

You can confirm that you are using the shared JavaScript runtime correctly by using the following instructions.

1. Open the **manifest.xml** file.
1. Find the `<Control xsi:type="Button" id="TaskpaneButton">` section and change the following `<Action ...>` XML.

    from:

    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```

    to:

    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```

1. Open the **./src/commands/commands.js** file.
1. Replace the **action** function with the code below. This will update the function to open and modify the task pane button to increment a counter. Opening and accessing the task pane DOM from a command only works with the shared JavaScript runtime.

    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

1. Save your changes and run the project.

   ```command line
   npm start
   ```

Each time you select the add-ins button, it will change the **run** button text to **go** and increment a counter after it.

## Runtime lifetime

When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`. Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.

> [!NOTE]
> The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running. Similarly, your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version. For more information, see [Runtimes](../reference/manifest/runtimes.md).

## About the shared JavaScript runtime

On Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments. This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.

However, you can configure your Office Add-in to share code in the same JavaScript runtime (also referred to as a shared runtime). This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.

Configuring a shared runtime enables the following scenarios.

- Your Office Add-in can use additional UI features:
  - [Add Custom keyboard shortcuts to your Office Add-ins (preview)](../design/keyboard-shortcuts.md)
  - [Create custom contextual tabs in Office Add-ins (preview)](../design/contextual-tabs.md)
  - [Enable and Disable Add-in Commands](../design/disable-add-in-commands.md)
  - [Run code in your Office Add-in when the document opens](run-code-on-document-open.md)
  - [Show or hide the task pane of your Office Add-in](show-hide-add-in.md)
- For Excel add-ins:
  - Custom functions will have full CORS support.
  - Custom functions can call Office.js APIs to read spreadsheet document data.

For Office on Windows, the shared runtime requires a Microsoft Internet Explorer 11 browser instance, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your add-in displays on the ribbon will run in the same shared runtime. The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.

![Diagram of a custom function, task pane, and ribbon buttons all running in a shared IE browser runtime in Excel.](../images/custom-functions-in-browser-runtime.png)

### Debugging

When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time. You'll need to use developer tools instead. For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

### Multiple task panes

Don't design your add-in to use multiple task panes if you are planning to use a shared runtime. A shared runtime only supports the use of one task pane. Note that any task pane without a `<TaskpaneID>` is considered a different task pane.

## Give us feedback

We'd love to hear your feedback on this feature. If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).

## See also

- [Call Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md)
- [Add custom keyboard shortcuts to your Office Add-ins (preview)](../design/keyboard-shortcuts.md)
- [Create custom contextual tabs in Office Add-ins (preview)](../design/contextual-tabs.md)
- [Enable and Disable Add-in Commands](../design/disable-add-in-commands.md)
- [Run code in your Office Add-in when the document opens](run-code-on-document-open.md)
- [Show or hide the task pane of your Office Add-in](show-hide-add-in.md)
- [Tutorial: Share data and events between Excel custom functions and the task pane](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
