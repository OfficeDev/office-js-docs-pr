---
title: Implement event-based activation in Excel, PowerPoint, and Word add-ins 
description: Learn how to develop an Excel, PowerPoint, and Word add-in that implements event-based activation.
ms.date: 06/30/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement event-based activation in Excel, PowerPoint, and Word add-ins 

Event-based activation automatically launches a centrally deployed Word, Excel, or PowerPoint add-in whenever a document is created or opened. This allows the add-in to validate, insert, or refresh critical content without any manual operations. The add-in is opened in the background to avoid disrupting the user.

> [!NOTE]
> For information on how to implement event-based activation for Outlook add-ins, see [Configure your Outlook add-in for event-based activation](../outlook/autolaunch.md).

## Supported events and clients

| Event name | Description | Supported clients and channels |
| ----- | ----- | ----- |
| `OnDocumentOpened` | Occurs when a user opens a document or creates a new document, spreadsheet, or presentation. | <ul><li>Windows (build >= 16.0.18324.20032)</li><li>Office on the web</li><li>Office on Mac will be available later </li></ul>|

## Behavior and limitations

As you develop an event-based add-in, be mindful of the following feature behaviors and limitations.

- Office on Mac on is not supported.
- The unified manifest is not supported.
- Event-based add-ins work only when deployed by an administrator. If users install them directly from AppSource or the Office Store, they will not automatically launch. Admin deployments are done by uploading the manifest to the Microsoft 365 admin center.
- If a user installs multiple add-ins that handle the same activation event, only one add-in will be activated. There is no deterministic way to know which add-in will be activated. For example, if multiple add-ins that handle `OnDocumentOpened`, only one of those handlers will run.
- APIs that interact with the UI or display UI elements are not supported for Word, PowerPoint, and Excel on Windows. This is because the event handler runs in a JavaScript-only runtime. For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md).

## Walkthrough: Automatically act when the document opens

The following sections walk you through how to develop a Word add-in that automatically changes the document header when a new or existing document opens. While this specific sample is for Word, the manifest configuration is the same for Excel and PowerPoint.

> [!IMPORTANT]
> This sample requires you to have a Microsoft 365 subscription with the supported version of Word.

### Create a new add-in

Create a new add-in by following the [Word add-in quick start](../quickstarts/word-quickstart-yo.md?tabs=yeoman). This will give you a working Office Add-in to which you can add the event-based activation code.

### Configure the manifest

To enable an event-based add-in, you must configure the following elements in the `VersionOverridesV1_0` node of the manifest.

- In the [Runtimes](/javascript/api/manifest/runtimes) element, make a new [Override element for Runtime](/javascript/api/manifest/override#override-element-for-runtime). Override the "javascript" type and reference the JavaScript file containing the function you want to trigger with the event.
- In the [DesktopFormFactor](/javascript/api/manifest/desktopformfactor) element, add a [FunctionFile](/javascript/api/manifest/functionfile) element for the JavaScript file with the event handler.
- In the [ExtensionPoint](/javascript/api/manifest/extensionpoint) element, set the `xsi:type` to `LaunchEvent`. This enables the event-based activation feature in your add-in.
- In the [LaunchEvent](/javascript/api/manifest/launchevent) element, set the `Type` to `OnDocumentOpened` and specify the JavaScript function name of the event handler in the `FunctionName` attribute.

Use the following sample manifest code to update your project.

1. In your code editor, open the quick start project you created.
1. Open the **manifest.xml** file located at the root of your project.
1. Select the entire **\<VersionOverrides\>** node (including the open and close tags) and replace it with the following XML.

    ```xml
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
        <Hosts>
          <Host xsi:type="Document">
            <Runtimes>
              <Runtime resid="Taskpane.Url" lifetime="long" />
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
            <bt:Url id="JsRuntimeWord.Url" DefaultValue="https://localhost:3000/commands.js"/>
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
            <bt:String id="CommandsGroup.Label" DefaultValue="Event-based add-in activation"/>
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

### Implement the event handler

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
        // If needed, Office.js is ready to be called
      });
      
      async function changeHeader(event) {
        Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
  
          if (body.text.length == 0) {
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
          if (results.items.length == 0) {
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
        Word.run(async (context) => {
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

### Test and validate your add-in

1. Run `npm start` to build your project and launch the web server. **Ignore the Word document that is opened**.
1. Manually sideload your add-in in Word on the web by following the guidance at [Sideload Office Add-ins to Office on the web](../testing/sideload-office-add-ins-for-testing.md#manually-sideload-an-add-in-to-office-on-the-web). Use the **manifest.xml** in the root of the project.
1. Try opening both new and existing Word documents in Word on the web. Headers should automatically be added when they open.

## Deploy your add-in

Event-based add-ins work only when deployed by an administrator. If users install them directly from AppSource or the Office Store, they will not automatically launch. To perform an admin deployment, upload the manifest to the Microsoft 365 admin center by taking the following actions.

1. In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.
1. On the **Integrated apps** page, choose the **Upload custom apps** action.

For more information about how to deploy an add-in, please refer to [Deploy and publish Office Add-ins in the Microsoft 365 admin center](/microsoft-365/admin/manage/office-addins).
