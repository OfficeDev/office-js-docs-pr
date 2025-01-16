---
title: Implement event-based activation in WXP add-ins (preview)
description: Learn how to develop a WXP add-in that implements event-based activation.
ms.date: 08/01/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement event-based activation in WXP add-ins (preview)

With the feature, you can develop an add-in to automatically activate and complete operations when certain events occur in Word, Excel, and PowerPoint, such as create new documents and open existing documents. 

> [!NOTE]
> The feature is in preview for early testing. You can not deploy a add-in with this feature to your customer yet. Please also notice the preview version of feature may be different from the released version.


## Supported events and clients

| Event name | Description | Supported clients and channels |
| ----- | ----- | ----- |
| `OnDocumentOpen` | Occurs on a user opens a document or creates a new document in WXP.| <ul><li> Office Win32 Desktop DevMain channel insider ring, version>= 16.0.18324.20032 </li></ul><ul><li> Office online </li></ul>|

The following sections walk you through how to develop a Word add-in that automatically changes the document header when a new or existing document opens. This highlights a sample scenario of how you can implement event-based activation in WXP add-ins.

## Set up your environment

To run the feature, you must have a supported version of Word and a Microsoft 365 subscription. Then, create a Word add-in project. You can create an add-in by following [Word add-in quick start](../quickstarts/word-quickstart-yo.md) and try to create an Office Add-in Task Pane project or other.

## Configure the manifest

 Currently, only Add-in only manifest supported. To enable an event-based add-in in WXP, you must configure the following elements in the `VersionOverridesV1_0` node of the manifest.

- In the [Runtimes](/javascript/api/manifest/runtimes) element, override the using runtime with a javascript type and reference a javascript file containing the function you want to execute.
- Set the `xsi:type` of the [ExtensionPoint](/javascript/api/manifest/extensionpoint) element to [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent). This enables the event-based activation feature in your WXP add-in.
- In the [LaunchEvent](/javascript/api/manifest/launchevent) element, set the `Type` to `OnDocumentOpen` and specify the JavaScript function name of the event handler in the `FunctionName` attribute.

### Code sample

1. In your code editor, open the quick start project you created.
1. Open the **manifest.xml** file located at the root of your project.
1. Select the entire **\<VersionOverrides\>** node (including the open and close tags) and replace it with the following XML. (Word version)

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
              <LaunchEvent Type="OnDocumentOpen" FunctionName="changeHeader"></LaunchEvent>
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
        <bt:Url id="JsRuntimeWord.Url" DefaultValue="https://raw.githubusercontent.com/yilin4/AddinForDLP/refs/heads/main/src/commands/autoruncommandsWord.js"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="Autorun For DLP"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="DLP"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded successfully. Go to the HOME tab and click the 'Show Task Pane' button to get started."/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane"/>
      </bt:LongStrings>
    </Resources>
</VersionOverrides>
```

4. Save your changes.

## Implement the event handler

To enable your add-in to complete tasks when the `OnDocumentOpen` event occurs, you must implement a JavaScript event handler. In this section, you'll create the `changeHeader` function that adds header of public or high confidential to a document when open it according to whether it's a new document or an old one that already has content.

1. From the same quick start project, navigate to the **./src/commands** directory.
1. In the **./src/commands** folder, create a new file named **autoruncommandsWord.js**.
1. Open the **autoruncommandsWord.js** file you created and add the following JavaScript code.

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
    if (body.text.length == 0)
    {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
      header.clear();
      firstPageHeader.clear();
      header.insertParagraph("Public - The data is for the public and shareable externally", "Start");
      firstPageHeader.insertParagraph("Public - The data is for the public and shareable externally", "Start");
      header.font.color = "#07641d";
      firstPageHeader.font.color = "#07641d";

      await context.sync();
    }
    else
    {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
      header.clear();
      firstPageHeader.clear();
      header.insertParagraph("High Confidential - The data must be secret or in some way highly critical", "Start");
      firstPageHeader.insertParagraph("High Confidential - The data must be secret or in some way highly critical", "Start");
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
    }
    else {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      header.clear();
      header.insertParagraph("High Confidential - The data must be secret or in some way highly critical", "Start");
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

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope

Office.actions.associate("changeHeader", changeHeader);
```

4. Save your changes. In the manifest, replace the following content to your own url.
```xml
<bt:Url id="JsRuntimeWord.Url" DefaultValue="https://raw.githubusercontent.com/yilin4/AddinForDLP/refs/heads/main/src/commands/autoruncommandsWord.js"/>
```

## Add a reference to the event-handling JavaScript file

Ensure that the **autoruncommandsWord.js** file must be a javascript file not a typescript file, and the online url is recommended.

## Test and validate your add-in

1. Sideload your add-in in [Word online](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) or Windows desktop.
1. Open a Word document and you will see the headers are added to the document.

## Behavior and limitations

As you develop an event-based add-in for WXP, be mindful of the following feature behaviors and limitations.
- Currently, the feature is only supported in add-in only manifest.
- Office MAC Desktop is not supported yet.
- If a user installs multiple add-ins with the same ativation event, only one add-in will be activated randomly.
- APIs that are not supported: [To be added]
