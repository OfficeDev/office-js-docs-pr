---
title: Add headers when a document opens
description: Learn how to develop a Word add-in that implements event-based activation to add headers when a document is opened.
ms.date: 09/02/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Add headers when a document opens

The following sections walk you through how to develop a Word add-in that automatically changes the document header when a new or existing document opens. While this specific add-in is for Word, the configuration of the manifest and the **webpack.config.js** file is the same for Excel and PowerPoint. For an overview of this event-based activation pattern, see [Activate add-ins with events](../develop/event-based-activation.md).

## Create a new add-in

Create a new add-in by following the [Word add-in quick start](../quickstarts/word-quickstart-yo.md?tabs=yeoman), but note the following changes in the steps described there.

- Use the guidance for the add-in only manifest. The unified manifest for Microsoft 365 doesn't yet support the `OnDocumentOpened` event that is used in this project.
- When Yo Office prompts you to choose a language, select *JavaScript*.

> [!NOTE]
> For a completed version of the sample described in this walkthrough, see the [Automatically add labels with an add-in when a Word document opens sample in our samples GitHub repo](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-label-on-open).

## Configure the manifest

To enable an event-based add-in, you must configure the following elements in the `VersionOverridesV1_0` node of the manifest. Note the following about the new markup which is provided below.

- The code that handles the `OnDocumentOpened` event runs in a browser runtime in Word on the web, but in a JavaScript-only runtime in Word on Windows. To configure this pattern, a [Runtime](/javascript/api/manifest/runtime) element is added that points browser runtimes to the project's **commands.html** file. This element has a child [Override element for Runtime](/javascript/api/manifest/override#override-element-for-runtime) that overrides the "javascript" type and points JavaScript-only runtimes to the **commands.js** file. For more information about runtimes for Office Add-ins, see [Runtimes in Office Add-ins](../testing/runtimes.md).
- In the [ExtensionPoint](/javascript/api/manifest/extensionpoint) element, the `xsi:type` is set to `LaunchEvent`. This enables the event-based activation feature in your add-in.
- In the [SourceLocation](/javascript/api/manifest/customfunctionssourcelocation) element of the **\<ExtensionPoint\>** element, the `resid` value is set to match that of the [Runtime](/javascript/api/manifest/runtime) element that references the HTML file.
- In the [LaunchEvent](/javascript/api/manifest/launchevent) element, the `Type` is set to `OnDocumentOpened` and the `FunctionName` attribute is set to the JavaScript function name of the event handler.
- In the `Resources` section, the `JsRuntimeWord.Url` is set to a **\public** subfolder in the web application. In conjunction with the changes you will make in the **webpack.config.js** file, this URL ensures that the **commands.js** that runs in the JavaScript-only runtime isn't bundled with code that requires a browser runtime. See [Configure webpack.config.js](#configure-webpackconfigjs).

Use the following sample manifest code to update your project.

1. In your code editor, open the quick start project you created.
1. Open the **manifest.xml** file located at the root of your project.
1. Select the entire `<VersionOverrides>` node (including the open and close tags) and replace it with the following XML.

    ```xml
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
        <Hosts>
          <Host xsi:type="Document">
            <Runtimes>
              <Runtime resid="WebViewRuntime.Url">
                <Override type="javascript" resid="JsRuntimeWord.Url"/>
              </Runtime>
            </Runtimes>
            <DesktopFormFactor>
              <GetStarted>
                <Title resid="GetStarted.Title"/>
                <Description resid="GetStarted.Description"/>
                <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
              </GetStarted>
              <FunctionFile resid="Commands.Url"/>
              <ExtensionPoint xsi:type="LaunchEvent">
                <LaunchEvents>
                  <LaunchEvent Type="OnDocumentOpened" FunctionName="changeHeader"></LaunchEvent>
                </LaunchEvents>
                <SourceLocation resid="WebViewRuntime.Url"/>
              </ExtensionPoint>
              <ExtensionPoint xsi:type="PrimaryCommandSurface">
                <OfficeTab id="TabHome">
                  <Group id="CommandsGroup">
                    <Label resid="CommandsGroup.Label"/>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Control xsi:type="Button" id="TaskpaneButton">
                      <Label resid="TaskpaneButton.Label"/>
                      <Supertip>
                        <Title resid="TaskpaneButton.Label"/>
                        <Description resid="TaskpaneButton.Tooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Icon.16x16"/>
                        <bt:Image size="32" resid="Icon.32x32"/>
                        <bt:Image size="80" resid="Icon.80x80"/>
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>ButtonId1</TaskpaneId>
                        <SourceLocation resid="Taskpane.Url"/>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>
            </DesktopFormFactor>
          </Host>
        </Hosts>
        <Resources>
          <bt:Images>
            <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
            <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
            <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
          </bt:Images>
          <bt:Urls>
            <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
            <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
            <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
            <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html"/>
            <bt:Url id="JsRuntimeWord.Url" DefaultValue="https://localhost:3000/public/commands.js"/>
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
            <bt:String id="CommandsGroup.Label" DefaultValue="Event-activated add-in"/>
            <bt:String id="TaskpaneButton.Label" DefaultValue="My add-in"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded successfully. Go to the HOME tab and click the 'Show Task Pane' button to get started."/>
            <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to show the task pane"/>
          </bt:LongStrings>
        </Resources>
      </VersionOverrides>
    ```

1. Save your changes.

## Implement the event handler

To enable your add-in to act when the `OnDocumentOpened` event occurs, you must implement a JavaScript event handler. In this section, you'll create the `changeHeader` function, which adds a "Public" header to new documents or a "Highly Confidential" header to existing documents that already have content.

1. In the **./src/commands** folder, open the file named **commands.js**.
1. Replace the entire contents of **commands.js** with the following JavaScript code.

    ```javascript
      /*
      * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
      * See LICENSE in the project root for license information.
      */
      /* global global, Office, self, window */
      
      Office.onReady(() => {
        // If needed, Office.js is ready to be called.
      });
      
      async function changeHeader(event) {
        await Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
  
          if (body.text.length === 0) {
          // For new or empty documents, make a "Public" header. 
            const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
            const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
            header.clear();
            firstPageHeader.clear();
  
            header.insertParagraph("Public - The data is for the public and shareable externally", "Start");
            firstPageHeader.insertParagraph("Public - The data is for the public and shareable externally", "Start");
            header.font.color = "#07641d";
            firstPageHeader.font.color = "#07641d";
            await context.sync();
          } else {
            // For existing documents, make a "Highly Confidential" header.
            const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
            const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
            header.clear();
            firstPageHeader.clear();
            header.insertParagraph("Highly Confidential - The data must be secret or in some way highly critical", "Start");
            firstPageHeader.insertParagraph("Highly Confidential - The data must be secret or in some way highly critical", "Start");
            header.font.color = "#f8334d";
            firstPageHeader.font.color = "#f8334d";
            await context.sync();
          }
        });
      
        // Calling event.completed is required. event.completed lets the platform know that processing has completed.
        event.completed();
      }
      
      async function paragraphChanged() {
        await Word.run(async (context) => {
          const results = context.document.body.search("110");
          results.load("length");
          await context.sync();
          if (results.items.length === 0) {
            const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
            header.clear();
            header.insertParagraph("Public - The data is for the public and shareable externally", "Start");
            const font = header.font;
            font.color = "#07641d";
      
            await context.sync();
          } else {
            const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
            header.clear();
            header.insertParagraph("Highly Confidential - The data must be secret or in some way highly critical", "Start");
            const font = header.font;
            font.color = "#f8334d";
  
            await context.sync();
          }
        });
      }
      
      async function registerOnParagraphChanged(event) {
        await Word.run(async (context) => {
          let eventContext = context.document.onParagraphChanged.add(paragraphChanged);
          await context.sync();
        });
        // Calling event.completed is required. event.completed lets the platform know that processing has completed.
        event.completed();
      }
            
      Office.actions.associate("changeHeader", changeHeader);
      Office.actions.associate("registerOnParagraphChanged", registerOnParagraphChanged);
    ```

1. Save your changes.

## Configure webpack.config.js

The **webpack.config.js** file needs to be configured so that it creates distinct bundles of the JavaScript code for browser and JavaScript-only runtimes. Take the following steps.

1. Add the following line to the top of the file where the other global `const`s are declared.

    ```
    const path = require("path");
    ```

1. To ensure that the add-in's icon can appear in the **Integrated Apps** list in the Microsoft 365 Admin portal, add the following property to the `devServer` object near the bottom of the file.

    ```
    allowedHosts: "all",
    ```

1. To ensure that the **commands.js** that runs in the JavaScript-only runtime isn't bundled with code that requires a browser runtime, add the following `static` property to the `devServer` object.

    ```
    static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/public",
      },
    ```

    The entire `devServer` object should now look like the following.

    ```
    devServer: {
      allowedHosts: "all",
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/public",
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
    ```

## Install the sample for testing

1. In a command prompt, navigate to the root of the project.
1. Run `npm run build:dev`.
1. Run `npm run dev-server`.
1. In the Microsoft 365 admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.
1. On the **Integrated apps** page, choose the **Upload custom apps** action.
1. On the **Upload Apps to deploy** page, select **Office Add-in** from the **App type** drop down.
1. In the **Choose how to upload app** section, select **Upload manifest file (.xml) from device**.
1. Use the file picker to navigate to the root of the project and then select the `manifest.xml` file.
1. Select **Just me** as the user.
1. Follow the instructions on screen to finish the deployment.

> [!IMPORTANT]
> You cannot run the add-in until after it has propagated to a platform. Propagation to Word on the web can take several hours, typically 2 to 3 hours. Propagation to Word on Windows can take 24 hours, typically 6 to 12 hours.
>
> To test whether the add-in has propagated, see [Try it out](#try-it-out).

## Try it out

1. In either Word on the web or Word on Windows, try opening both new and existing Word documents. If the add-in has propagated to the platform, headers should automatically be added when the document opens, and there should be a **My Add-in** button in an **Event-activated add-in** group on the **Home** tab of the ribbon. If these things don't happen, propagation to the platform hasn't completed. Close Word and try again in a while.
1. Select the **My add-ins** button to open the task pane.
1. Select any of the links on the task pane to add or change the header.

> [!IMPORTANT] 
> When you're finished working with the sample, [uninstall it](#uninstall-the-add-in).

## Uninstall the add-in

To uninstall the add-in, take the following steps:

1. In the Microsoft 365 admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.
1. On the **Integrated apps** page, select the add-in.
1. On the add-in's flyout, select **Remove app**.
1. On the **Remove apps** page, confirm that you want to remove the app and select **Remove**.
1. On the **Successfully removed** page, select **Done**.

> [!IMPORTANT] 
> Uninstallation must propagate to the platforms just as installation does. Propagation to Word on the web can take several hours, typically 2 to 3 hours. Propagation to Word on Windows can take 24 hours, typically 6 to 12 hours.
>
> To test if uninstallation has propagated, open a Word file on the platform. If the **My Add-in** button in an **Event-activated add-in** group is still on the **Home** tab of the ribbon, propagation hasn't happened. Close Word and try again in a while.

## See also

- [Activate add-ins with events](../develop/event-based-activation.md)
- [Debug event-based or spam-reporting add-ins](../testing/debug-autolaunch.md)
- [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
