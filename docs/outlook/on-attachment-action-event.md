---
title: Check an attachment before it's opened, saved, or downloaded (preview)
description: Learn how to implement an event-based Outlook add-in that checks an attachment before it's opened, saved, or downloaded.
ms.date: 05/14/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Check an attachment before it's opened, saved, or downloaded (preview)

Easily manage attachments securely before a user opens, saves, or downloads them from mail items. With the `OnAttachmentAction` event, automatically activate your add-in whenever a supported attachment action occurs in read mode to:

- Log and audit attachment interactions for compliance workflows.
- Verify that an attachment isn't a malicious file.
- Apply organization-specific policies to attachments.

The following sections demonstrate how to implement an event-based Outlook add-in that handles the `OnAttachmentAction` event. By the end of this walkthrough, you'll have an add-in that detects attachment actions and runs your custom checks and attachment-management operations.

> [!NOTE]
>
> The `OnAttachmentAction` event is in [preview](/javascript/api/requirement-sets/outlook/outlook-requirement-set-preview). Features in preview shouldn't be used in production add-ins as they may change based on feedback we receive. We invite you to try out this feature in test or development environments and welcome feedback on your experience through GitHub (see the "Office Add-ins feedback" section at the end of this page).

## Supported clients and modes

The `OnAttachmentAction` event is supported in the following clients in Message Read and Appointment Read modes.

|Client|Support status|
|-----|-----|
|**Web browser**|Supported|
|[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)|Supported|
|**Windows (classic)**|Not supported|
|**Mac**|Not supported|
|**Android**|Not supported|
|**iOS**|Not supported|

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart-yo.md), which creates an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

> [!NOTE]
>
> The `OnAttachmentAction` event is only supported in the add-in only manifest at this time. We're working to support the event in the unified manifest for Microsoft 365. To learn more about manifest types, see [Office Add-ins manifest](../develop/add-in-manifests.md).

1. In your add-in project, open the **manifest.xml** file.

1. Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML. Note the following about the code.
    - The [\<Runtimes>](/javascript/api/manifest/runtimes) and [\<LaunchEvent>](/javascript/api/manifest/extensionpoint#launchevent) elements must be configured in the nested `VersionOverridesV1_1` node.
    - Outlook on the web and the new Outlook on Windows run the event handlers in a browser runtime. In the `resid` attribute of the [\<Runtime>](/javascript/api/manifest/runtime) element, specify the HTML file that contains or references the JavaScript event handlers. For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md).
    - To enable event-based activation, set the `xsi:type` attribute of the [\<ExtensionPoint>](/javascript/api/manifest/extensionpoint#launchevent) element to `LaunchEvent`.
    - To activate the add-in when an attachment action occurs, set the `Type` of the child [\<LaunchEvent>](/javascript/api/manifest/launchevent) element to `OnAttachmentAction` and specify the name of the event handler in the `FunctionName` attribute.
    - Specify the HTML file that contains or references the event handlers in the child [\<SourceLocation>](/javascript/api/manifest/customfunctionssourcelocation) element. Its `resid` value must match the `resid` of the browser runtime declared in the \<Runtime> element.

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.16">
            <bt:Set Name="Mailbox" />
          </bt:Sets>
        </Requirements>
        <Hosts>
          <Host xsi:type="MailHost">
            <!-- Event-based activation happens in a lightweight runtime.-->
            <Runtimes>
              <!-- HTML file including reference to or inline JavaScript event handlers.
                   This is used by Outlook on the web and the new Outlook on Windows. -->
              <Runtime resid="WebViewRuntime.Url" />
            </Runtimes>
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageReadCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgReadGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgReadOpenPaneButton">
                      <Label resid="TaskpaneButton.Label" />
                      <Supertip>
                        <Title resid="TaskpaneButton.Label" />
                        <Description resid="TaskpaneButton.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Icon.16x16" />
                        <bt:Image size="32" resid="Icon.32x32" />
                        <bt:Image size="80" resid="Icon.80x80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <SourceLocation resid="Taskpane.Url" />
                      </Action>
                    </Control>
                    <Control xsi:type="Button" id="ActionButton">
                      <Label resid="ActionButton.Label"/>
                      <Supertip>
                        <Title resid="ActionButton.Label"/>
                        <Description resid="ActionButton.Tooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Icon.16x16"/>
                        <bt:Image size="32" resid="Icon.32x32"/>
                        <bt:Image size="80" resid="Icon.80x80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>action</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>
              <!-- Launches the add-in when an attachment action is performed. -->
              <ExtensionPoint xsi:type="LaunchEvent">
                <LaunchEvents>
                  <LaunchEvent Type="OnAttachmentAction" FunctionName="onAttachmentActionHandler"/>
                </LaunchEvents>
                <!-- Identifies the runtime to be used. The resid value must match that of the Runtime element that represents the browser runtime. -->
                <SourceLocation resid="WebViewRuntime.Url"/>
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
            <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
            <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
            <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
            <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
            <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
            <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
          </bt:LongStrings>
        </Resources>
      </VersionOverrides>
    </VersionOverrides>
    ```

1. Save your changes.

> [!TIP]
> To learn about runtimes in add-ins, see [Runtimes in Office Add-ins](../testing/runtimes.md).

## Implement the event handler

The event handler specifies the operations you want your add-in to run when the `OnAttachmentAction` event occurs. The following steps show how to implement a sample handler for the `OnAttachmentAction` event.

1. Go to the **./src** directory of the project you created. Then, create a new folder named **launchevent**.
1. In the **./src/launchevent** folder, create a new file named **launchevent.js**.
1. In the **launchevent.js** file, add the following JavaScript code. Note the following about the code.
  
    - The [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) method must be called from your event-handling code. This ensures that the function name specified in the \<LaunchEvent> element of your manifest is mapped to its respective JavaScript counterpart.
    - The `event` object that the handler receives includes information about the attachment, such as its identifier. For information about the attachment details, see [Office.AttachmentActionEventArgs](/javascript/api/outlook/office.attachmentactioneventargs?view=outlook-js-preview&preserve-view=true).
    - To signal that the add-in has completed processing an attachment, your event handler must call the [event.completed](/javascript/api/outlook/office.mailboxevent#outlook-office-mailboxevent-completed-member(1)) method. Note that code included after the `event.completed` method isn't guaranteed to run.

    ```javascript
    function onAttachmentActionHandler(event) {
      const attachmentIds = event.attachmentIds;

      // Log details for diagnostics.
      console.log(`attachment IDs: ${attachmentIds.join(", ")}`);
      console.log(`Number of attachments in event: ${attachmentIds.length}`);

      // Perform other operations on attachments here.

      event.completed();
    }

    // Map the manifest handler name to its JavaScript implementation.
    Office.actions.associate("onAttachmentActionHandler", onAttachmentActionHandler);
    ```

1. Save your changes.

## Update the commands HTML file

1. From the **./src/commands** folder, open **commands.html**.

1. Add the following code below the existing **script** tag.

   ```html
   <script type="text/javascript" src="../launchevent/launchevent.js"></script>
   ```

1. Save your changes.

## Update webpack config settings

1. From the root directory of the project, open the **webpack.config.js** file.

1. Locate the `plugins` array within the `config` object and add the following new object to the beginning of the array.

    ```js
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "./src/launchevent/launchevent.js",
          to: "launchevent.js",
        },
      ],
    }),
    ```

1. Save your changes.

## Try it out

1. Run the following commands in the root directory of your project. When you run `npm start`, the local web server will start (if it isn't already running) and your add-in will be sideloaded.

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    [!INCLUDE [outlook-manual-sideloading](../includes/outlook-manual-sideloading.md)]

1. In a supported Outlook client, open an existing message with a file or Outlook mail item attachment.

    > [!TIP]
    > If you don't have a message with an attachment in your inbox, send yourself a test message with a file attachment or Outlook mail item.

1. Open Microsoft Edge Developer Tools.

    - **Outlook on the web**: Select <kbd>F12</kbd> or select and hold (or right-click) anywhere in the message, then select **Inspect**. Open the **Console** tab.
    - **New Outlook on Windows**: Follow the instructions in the "Debug your add-in" section of [Develop Outlook add-ins for the new Outlook on Windows](one-outlook.md#debug-your-add-in). Then, open the **Console** tab.
1. Select the attachment to open it. Alternatively, select another supported attachment action, such as **Copy**, **Save**, or **Download**.

    The `OnAttachmentAction` event occurs and details about the attachment are logged to the console.

1. [!INCLUDE [Instructions to stop web server and uninstall dev add-in](../includes/stop-uninstall-outlook-dev-add-in.md)]

## Event behavior and limitations

Because the `OnAttachmentAction` event is part of the event-based activation feature, the same behaviors and limitations apply. For a detailed description, see [Activate add-ins with events](../develop/event-based-activation.md#behavior-and-limitations). Additionally, be mindful of the following behaviors and constraints when implementing the `OnAttachmentAction` event.

- The `OnAttachmentAction` event occurs on a message or appointment in read mode when an attachment is opened, copied, saved, or downloaded. The following actions aren't supported.
  - Dragging and dropping an attachment from a mail item.
  - Selecting the **View in OneDrive** or **Upload to OneDrive** option.
  - Previewing an attachment.

    > [!NOTE]
    > In Outlook on the web, double-clicking an attachment opens the attachment in Preview mode.

- The `OnAttachmentAction` event occurs on the following attachment types.
  - File attachments.
  - Outlook mail items.

  Inline and cloud attachments aren't supported.
- The `OnAttachmentAction` event only occurs on attachments contained in a message or appointment formatted in HTML. Rich Text Format (RTF) isn't currently supported.
- When the event handler runs, a notification alerts the user that an add-in is processing the attachment. A **Skip** action is available on the notification so that the user can stop the handler and immediately continue with their attachment action. For the `OnAttachmentAction` event, the handler can block the attachment action for up to 15 seconds.
- The attachment action proceeds when the event handler completes, is skipped, or times out.
- If a user attempts to download an attachment while the event handler is running, the download only proceeds after the handler operation completes.
- If the event-based add-in becomes unavailable (for example, there's an error preventing the add-in from loading), the user is still able to interact with the attachment.
- If a user switches to another attachment or mail item while the event-based add-in is processing an attachment, the add-in stops running.

## See also

- [Activate add-ins with events](../develop/event-based-activation.md)
- [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md)
