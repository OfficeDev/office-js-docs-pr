---
ms.date: 08/13/2020
title: "Configure your Excel add-in to share the browser runtime"
ms.prod: excel
description: 'Configure your Excel add-in to share the browser runtime and run ribbon, task pane, and custom function code in the same runtime.'
localization_priority: Priority
---

# Configure your Excel add-in to use a shared JavaScript runtime

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments. This creates limitations such as not being able to easily share global data, and not having access to all CORS functionality from a custom function.

However, you can configure your Excel add-in to share code in a shared JavaScript runtime. This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in. It also enables you to run code when the document opens, or to run code while the task pane is closed. To configure your add-in to use a shared runtime, follow the instructions in this article.

## Create the add-in project

If you are starting a new project, follow these steps to use the Yeoman generator to create an Excel add-in project. Run the following command and then answer the prompts with the following answers:

```command line
yo office
```

- Choose a project type: **Excel Custom Functions Add-in project**
- Choose a script type: **JavaScript**
- What do you want to name your add-in? **My Office Add-in**

![Screenshot of answering prompts from yo office to create the add-in project.](../images/yo-office-excel-project.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

## Configure the manifest

Follow these steps for a new or existing project to configure it to use a shared runtime.

1. Start Visual Studio Code and open the **My Office Add-in** project.
2. Open the **manifest.xml** file.
3. Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section. The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed. The resid is `ContosoAddin.Url` which references a string in the resources section later. You can use any resid value you want, but it should match the resid of the other elements in your add-in elements.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**. This resid matches the `<Runtime>` resid element. Note that if you don't have custom functions, you will not have a **Page** entry and can skip this step.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**. Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**. Note that if you don't have a task pane, you won't have a **ShowTaskpane** action, and can skip this step.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/dist/taskpane.html"/>
   ...
   ```

8. Make sure the taskpane.html has a `<script>` tag that references the dist/functions.js file. The following is an example:

   ```html
   <script type=”text/javascript” src=”/dist/functions.js” ></script>
   ```

   > [!NOTE]
   > If the add-in uses Webpack and the HtmlWebpackPlugin to insert script tags, as add-ins created by the Yeoman generator do (see [Create the add-in project](#create-the-add-in-project) above), then you must ensure that the functions.js module is included in the `chunks` array as in the following example:
   >
   > ```javascript
   > new HtmlWebpackPlugin({
   >     filename: "taskpane.html",
   >     template: "./src/taskpane/taskpane.html",
   >     chunks: ["polyfill", "taskpane", “functions”]
   > }),
   >```

9. Save your changes and rebuild the project.

   ```command line
   npm run build
   ```

## Runtime lifetime

When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`. Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.

>[!NOTE]
> The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running. Similarly your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

>[!NOTE]
> If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version. For more information, see [Runtimes](../reference/manifest/runtimes.md).

## Multiple task panes

Don't design your add-in to use multiple task panes if you are planning to use a shared runtime. A shared runtime only supports the use of one task pane. Note that any task pane without a `<TaskpaneID>` is considered a different task pane.

## Next steps

- Read the [Call Excel APIs from a custom function](call-excel-apis-from-custom-function.md) article for details on using the Excel JavaScript APIs and custom Excel functions in a shared runtime.
- Explore the patterns-and-practices sample [Manage ribbon and task pane UI, and run code on doc open](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) to see a larger example of the shared JavaScript runtime in action.

## See also

- [Overview: Run your add-in code in a shared JavaScript runtime](custom-functions-shared-overview.md)
