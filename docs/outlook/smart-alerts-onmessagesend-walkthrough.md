---
title: Use Smart Alerts and the OnMessageSend and OnAppointmentSend events in your Outlook add-in
description: Learn how to handle the on-send events in your Outlook add-in using event-based activation.
ms.topic: article
ms.date: 09/07/2022
ms.localizationpriority: medium
---

# Use Smart Alerts and the OnMessageSend and OnAppointmentSend events in your Outlook add-in

The `OnMessageSend` and `OnAppointmentSend` events take advantage of Smart Alerts, which allows you to run logic after a user selects **Send** in their Outlook message or appointment. Your event handler allows you to give your users the opportunity to improve their emails and meeting invites before they're sent.

The following walkthrough uses the `OnMessageSend` event. By the end of this walkthrough, you'll have an add-in that runs whenever a message is being sent and checks if the user forgot to add a document or picture they mentioned in their email.

> [!NOTE]
> The `OnMessageSend` and `OnAppointmentSend` events were introduced in [requirement set 1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12). See [clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) that support this requirement set.

## Prerequisites

The `OnMessageSend` event is available through the event-based activation feature. To understand how to configure your add-in to use this feature, use other available events, debug your add-in, and more, see [Configure your Outlook add-in for event-based activation](autolaunch.md).

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator), which creates an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

1. In your code editor, open the quick start project.

1. Open the **manifest.xml** file located at the root of your project.

1. Select the entire **\<VersionOverrides\>** node (including open and close tags) and replace it with the following XML, then save your changes.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.12">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and on the new Mac UI. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook on Windows. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
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

          <!-- Can configure other command surface extension points for add-in command support. -->

          <!-- Enable launching the add-in on the included event. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
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
        <!-- Entry needed for Outlook on Windows. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/launchevent.js" />
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

> [!TIP]
>
> - For **SendMode** options available with the `OnMessageSend` and `OnAppointmentSend` events, see [Available SendMode options](/javascript/api/manifest/launchevent#available-sendmode-options).
> - To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).

## Implement event handling

You have to implement handling for your selected event.

In this scenario, you'll add handling for sending a message. Your add-in will check for certain keywords in the message. If any of those keywords are found, it will then check if there are any attachments. If there are no attachments, your add-in will recommend to the user to add the possibly missing attachment.

1. From the same quick start project, create a new folder named **launchevent** under the **./src** directory.

1. In the **./src/launchevent** folder, create a new file named **launchevent.js**.

1. Open the file **./src/launchevent/launchevent.js** in your code editor and add the following JavaScript code.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { asyncContext: event },
        getBodyCallback
      );
    }

    function getBodyCallback(asyncResult){
      let event = asyncResult.asyncContext;
      let body = "";
      if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
        body = asyncResult.value;
      } else {
        let message = "Failed to get body text";
        console.error(message);
        event.completed({ allowEvent: false, errorMessage: message });
        return;
      }

      let matches = hasMatches(body);
      if (matches) {
        Office.context.mailbox.item.getAttachmentsAsync(
          { asyncContext: event },
          getAttachmentsCallback);
      } else {
        event.completed({ allowEvent: true });
      }
    }

    function hasMatches(body) {
      if (body == null || body == "") {
        return false;
      }

      const arrayOfTerms = ["send", "picture", "document", "attachment"];
      for (let index = 0; index < arrayOfTerms.length; index++) {
        const term = arrayOfTerms[index].trim();
        const regex = RegExp(term, 'i');
        if (regex.test(body)) {
          return true;
        }
      }

      return false;
    }

    function getAttachmentsCallback(asyncResult) {
      let event = asyncResult.asyncContext;
      if (asyncResult.value.length > 0) {
        for (let i = 0; i < asyncResult.value.length; i++) {
          if (asyncResult.value[i].isInline == false) {
            event.completed({ allowEvent: true });
            return;
          }
        }

        event.completed({ allowEvent: false, errorMessage: "Looks like you forgot to include an attachment?" });
      } else {
        event.completed({ allowEvent: false, errorMessage: "Looks like you're forgetting to include an attachment?" });
      }
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

## Update the commands HTML file

1. In the **./src/commands** folder, open **commands.html**.

1. Immediately before the closing **head** tag (`</head>`), add a script entry for the event-handling JavaScript code.

   ```js
   <script type="text/javascript" src="../launchevent/launchevent.js"></script> 
   ```

1. Save your changes.

## Update webpack config settings

1. Open the **webpack.config.js** file found in the root directory of the project and complete the following steps.

1. Locate the `plugins` array within the `config` object and add this new object to the beginning of the array.

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

1. Run the following commands in the root directory of your project. When you run `npm start`, the local web server will start (if it's not already running) and your add-in will be sideloaded.

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.

1. In Outlook on Windows, create a new message and set the subject. In the body, add text like "Hey, check out this picture of my dog!".
1. Send the message. A dialog should pop up with a recommendation for you to add an attachment.

    ![Dialog recommending that the user include an attachment.](../images/outlook-win-smart-alert.png)

1. Add an attachment then send the message again. There should be no alert this time.

## Deploy to users

Similar to other event-based add-ins, add-ins that use the Smart Alerts feature must be deployed by an organization's administrator. For guidance on how to deploy your add-in via the Microsoft 365 admin center, see the **Deploy to users** section in [Configure your Outlook add-in for event-based activation](autolaunch.md#deploy-to-users).

> [!IMPORTANT]
> Add-ins that use the Smart Alerts feature can only be published to AppSource if the manifest's [SendMode property](/javascript/api/manifest/launchevent) is set to the `SoftBlock` or `PromptUser` option. If an add-in's **SendMode** property is set to `Block`, it can only be deployed by an organization's admin as it will fail AppSource validation. To learn more about publishing your event-based add-in to AppSource, see [AppSource listing options for your event-based Outlook add-in](autolaunch-store-options.md).

## Smart Alerts feature behavior and scenarios

Descriptions of the **SendMode** options and recommendations for when to use them are detailed in [Available SendMode options](/javascript/api/manifest/launchevent#available-sendmode-options). The following describes the feature's behavior for certain scenarios.

### Add-in is unavailable

If the add-in is unavailable when a message or appointment is being sent (for example, an error occurs that prevents the add-in from loading), the user is alerted. The options available to the user differ depending on the **SendMode** option applied to the add-in.

If the `PromptUser` or `SoftBlock` option is used, the user can choose **Send Anyway** to send the item without the add-in checking it, or **Try Later** to let the item be checked by the add-in when it becomes available again.

![Dialog that alerts the user that the add-in is unavailable and gives the user the option to send the item now or later.](../images/outlook-soft-block-promptUser-unavailable.png)

If the `Block` option is used, the user can't send the item until the add-in becomes available.

![Dialog that alerts the user that the add-in is unavailable. The user can only send the item when the add-in is available again.](../images/outlook-hard-block-unavailable.png)

### Long-running add-in operations

If the add-in runs for more than five seconds, but less than five minutes, the user is alerted that the add-in is taking longer than expected to process the message or appointment.

If the `PromptUser` option is used, the user can choose **Send Anyway** to send the item without the add-in completing its check. Alternatively, the user can select **Don't Send** to stop the add-in from processing.

![Dialog that alerts the user that the add-in is taking longer than expected to process the item. The user can choose to send the item without the add-in completing its check or stop the add-in from processing the item.](../images/outlook-promptUser-long-running.png)

However, if the `SoftBlock` or `Block` option is used, the user will not be able to send the item until the add-in completes processing it.

![Dialog that alerts the user that the add-in is taking longer than expected to process the item. The user must wait until the add-in completes processing the item before it can be sent.](../images/outlook-soft-hard-block-long-running.png)

`OnMessageSend` and `OnAppointmentSend` add-ins should be short-running and lightweight. To avoid the long-running operation dialog, use other events to process conditional checks before the `OnMessageSend` or `OnAppointmentSend` event is activated. For example, if the user is required to encrypt attachments for every message or appointment, consider using the `OnMessageAttachmentsChanged` or `OnAppointmentAttachmentsChanged` event to perform the check.

### Add-in timed out

If the add-in runs for five minutes or more, it will time out. If the `PromptUser` option is used, the user can choose **Send Anyway** to send the item without the add-in completing its check. Alternatively, the user can choose **Don't Send**.

![Dialog that alerts the user that the add-in process has timed out. The user can choose to send the item without the add-in completing its check, or not send the item.](../images/outlook-promptUser-timeout.png)

If the `SoftBlock` or `Block` option is  used, the user can't send the item until the add-in completes its check. The user must attempt to send the item again to reactivate the add-in.

![Dialog that alerts the user that the add-in process has timed out. The user must attempt to send the item again to activate the add-in before they can send the message or appointment.](../images/outlook-soft-hard-block-timeout.png)

## Limitations

Because the `OnMessageSend` and `OnAppointmentSend` events are supported through the event-based activation feature, the same feature limitations apply to add-ins that activate as a result of these events. For a description of these limitations, see [Event-based activation behavior and limitations](autolaunch.md#event-based-activation-behavior-and-limitations).

In addition to these constraints, only one instance each of the `OnMessageSend` and `OnAppointmentSend` event can be declared in the manifest. If you require multiple `OnMessageSend` or `OnAppointmentSend` events, you must declare each one in a separate manifest or add-in.

While a Smart Alerts dialog message can be changed to suit your add-in scenario using the [errorMessage property](/javascript/api/office/office.addincommands.eventcompletedoptions) of the event.completed method, the following can't be customized.

- The dialog's title bar. Your add-in's name is always displayed there.
- The message's format. For example, you can't change the text's font size and color or insert a bulleted list.
- The dialog options. For example, the **Send Anyway** and **Don't Send** options are fixed and depend on the [SendMode option](/javascript/api/manifest/launchevent#available-sendmode-options) you select.
- Event-based activation processing and progress information dialogs. For example, the text and options that appear in the timeout and long-running operation dialogs can't be changed.

## See also

- [Outlook add-in manifests](manifests.md)
- [Configure your Outlook add-in for event-based activation](autolaunch.md)
- [How to debug event-based add-ins](debug-autolaunch.md)
- [AppSource listing options for your event-based Outlook add-in](autolaunch-store-options.md)
- [Office Add-ins code sample: Use Outlook Smart Alerts](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
