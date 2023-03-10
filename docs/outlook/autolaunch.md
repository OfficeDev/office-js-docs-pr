---
title: Configure your Outlook add-in for event-based activation
description: Learn how to configure your Outlook add-in for event-based activation.
ms.topic: article
ms.date: 03/10/2023
ms.localizationpriority: medium
---

# Configure your Outlook add-in for event-based activation

Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks. This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item. You can also integrate with the task pane and function commands.

By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.

> [!NOTE]
> Support for this feature was introduced in [requirement set 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), with additional events now available in subsequent requirement sets. For details about an event's minimum requirement set and the clients and platforms that support it, see [Supported events](#supported-events) and [Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).
>
> Event-based activation isn't supported in Outlook on iOS or Android.

> [!IMPORTANT]
> Event-based activation isn't yet supported for the [Teams manifest for Office Add-ins (preview)](../develop/json-manifest-overview.md). We are working on providing that support soon.

## Supported events

The following table lists events that are currently available and the supported clients for each event. When an event is raised, the handler receives an `event` object which may include details specific to the type of event. The **Description** column includes a link to the related object where applicable.

|Event canonical name</br>and XML manifest name|Teams manifest name|Description|Minimum requirement set and supported clients|
|---|---|---|---|
|`OnNewMessageCompose`| newMessageComposeCreated |On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI |
|`OnNewAppointmentOrganizer`|newAppointmentOrganizerCreated|On creating a new appointment but not on editing an existing one.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI |
|`OnMessageAttachmentsChanged`|messageAttachmentsChanged|On adding or removing attachments while composing a message.<br><br>Event-specific data object: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnAppointmentAttachmentsChanged`|appointmentAttachmentsChanged|On adding or removing attachments while composing an appointment.<br><br>Event-specific data object: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnMessageRecipientsChanged`|messageRecipientsChanged|On adding or removing recipients while composing a message.<br><br>Event-specific data object: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnAppointmentAttendeesChanged`|appointmentAttendeesChanged|On adding or removing attendees while composing an appointment.<br><br>Event-specific data object: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnAppointmentTimeChanged`|appointmentTimeChanged|On changing date/time while composing an appointment.<br><br>Event-specific data object: [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnAppointmentRecurrenceChanged`|appointmentRecurrenceChanged|On adding, changing, or removing the recurrence details while composing an appointment. If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.<br><br>Event-specific data object: [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnInfoBarDismissClicked`|infoBarDismissClicked|On dismissing a notification while composing a message or appointment item. Only the add-in that added the notification will be notified.<br><br>Event-specific data object: [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnMessageSend`|messageSending|On sending a message item. To learn more, see the [Smart Alerts walkthrough](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnAppointmentSend`|appointmentSending|On sending an appointment item. To learn more, see the [Smart Alerts walkthrough](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnMessageCompose`|messageComposeOpened|On composing a new message (includes reply, reply all, and forward) or editing a draft.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnAppointmentOrganizer`|appointmentOrganizerOpened|On creating a new appointment or editing an existing one.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|
|`OnMessageFromChanged`|Not available|On changing the mail account in the **From** field of a message being composed. To learn more, see [Automatically update your signature when switching between Exchange accounts (preview)](onmessagefromchanged-onappointmentfromchanged-events.md).|[Preview](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview)<br><br>- Windows<sup>1</sup>|
|`OnSensitivityLabelChanged`|Not available|On changing the sensitivity label while composing a message or appointment. To learn how to manage the sensitivity label of a mail item, see [Manage the sensitivity label of your message or appointment in compose mode (preview)](sensitivity-label.md).<br><br>Event-specific data object: [SensitivityLabelChangedEventArgs](/javascript/api/outlook/office.sensitivitylabelchangedeventargs?view=outlook-js-preview&preserve-view=true)|[Preview](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview)<br><br>- Windows<sup>1</sup><br>- Web browser<br>- New Mac UI|

> [!NOTE]
> <sup>1</sup> Event-based add-ins in Outlook on Windows require a minimum of Windows 10 Version 1903 (Build 18362) or Windows Server 2019 Version 1903 to run.

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

To configure the manifest, select the tab for the type of manifest you are using.

# [XML Manifest](#tab/xmlmanifest)

To enable event-based activation of your add-in, you must configure the [Runtimes](/javascript/api/manifest/runtimes) element and [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) extension point in the `VersionOverridesV1_1` node of the manifest. For now, `DesktopFormFactor` is the only supported form factor.

1. In your code editor, open the quick start project.

1. Open the **manifest.xml** file located at the root of your project.

1. Select the entire **\<VersionOverrides\>** node (including open and close tags) and replace it with the following XML, then save your changes.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.10">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and Outlook on the new Mac UI. -->
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

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewAppointmentComposeHandler"/>
              
              <!-- Other available events -->
              <!--
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnMessageCompose" FunctionName="onMessageComposeHandler" />
              <LaunchEvent Type="OnAppointmentOrganizer" FunctionName="onAppointmentOrganizerHandler" />
              -->
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

Outlook on Windows uses a JavaScript file, while Outlook on the web and on the new Mac UI use an HTML file that can reference the same JavaScript file. You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client. As such, to configure event handling, provide the location of the HTML in the **\<Runtime\>** element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.

# [Teams Manifest (developer preview)](#tab/jsonmanifest)

> [!IMPORTANT]
> Event based activation isn't yet supported for the [Teams manifest for Office Add-ins (preview)](../develop/json-manifest-overview.md). This tab is for future use.

1. Open the **manifest.json** file.

1. Add the following object to the "extensions.runtimes" array. Note the following about this markup:

   - The "minVersion" of the Mailbox requirement set is set to "1.10" because the table earlier in this article specifies that this is the lowest version of the requirement set that supports the `OnNewMessageCompose` and `OnNewAppointmentCompose` events.
   - The "id" of the runtime is set to the descriptive name "autorun_runtime".
   - The "code" property has a child "page" property that is set to an HTML file and a child "script" property that is set to a JavaScript file. You'll create or edit these files in later steps. Office uses one of these values depending on the platform.
       - Office on Windows executes the event handlers in a JavaScript-only runtime, which loads a JavaScript file directly.
       - Office on Mac and the web execute the handlers in a browser runtime, which loads an HTML file. That file, in turn, contains a `<script>` tag that loads the JavaScript file.
     For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md).
   - The "lifetime" property is set to "short", which means that the runtime starts up when one of the events is triggered and shuts down when the handler completes. (In certain rare cases, the runtime shuts down before the handler completes. See [Runtimes in Office Add-ins](../testing/runtimes.md).)
   - There are two types of "actions" that can run in the runtime. You'll create functions to correspond to these actions in a later step.

    ```json
     {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.10"
                }
            ]
        },
        "id": "autorun_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/launchevent.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "onNewMessageComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewMessageComposeHandler"
            },
            {
                "id": "onNewAppointmentComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewAppointmentComposeHandler"
            }
        ]
    }
    ```

1. Add the following "autoRunEvents" array as a property of the object in the "extensions" array.

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. Add the following object to the "autoRunEvents" array. The "events" property maps handlers to events as described in the table earlier in this article. The handler names must match those used in the "id" properties of the objects in the "actions" array in an earlier step.

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.10"
                  }
              ],
              "scopes": [
                  "mail"
              ]
          },
          "events": [
              {
                  "type": "newMessageComposeCreated",
                  "actionId": "onNewMessageComposeHandler"
              },
              {
                  "type": "newAppointmentOrganizerCreated",
                  "actionId": "onNewAppointmentComposeHandler"
              }
          ]
      }
    ```

---

> [!TIP]
>
> - To learn about runtimes in add-ins, see [Runtimes in Office Add-ins](../testing/runtimes.md).
> - To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).

## Implement event handling

You have to implement handling for your selected events.

In this scenario, you'll add handling for composing new items.

1. From the same quick start project, create a new folder named **launchevent** under the **./src** directory.

1. In the **./src/launchevent** folder, create a new file named **launchevent.js**.

1. Open the file **./src/launchevent/launchevent.js** in your code editor and add the following JavaScript code.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onNewMessageComposeHandler(event) {
      setSubject(event);
    }
    function onNewAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext": event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
          }

          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
        });
    }

    // IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
      Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
      Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
    }
    ```

1. Save your changes.

> [!IMPORTANT]
> Windows: At present, imports aren't supported in the JavaScript file where you implement the handling for event-based activation.

> [!TIP]
> Event-based add-ins running in Outlook on Windows don't run code included in the `Office.onReady()` and `Office.initialize` functions. We recommend adding your add-in startup logic, such as checking the user's Outlook version, to your event handlers instead.

## Update the commands HTML file

1. In the **./src/commands** folder, open **commands.html**.

1. Immediately before the closing **head** tag (`</head>`), add a script entry to include the event-handling JavaScript code.

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. Save your changes.

## Update webpack config settings

1. Open the **webpack.config.js** file found in the root directory of the project and complete the following steps.

1. Locate the `plugins` array within the `config` object and add this new object at the beginning of the array.

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

1. In Outlook on the web, create a new message.

    ![A message window in Outlook on the web with the subject set on compose.](../images/outlook-web-autolaunch-1.png)

1. In Outlook on the new Mac UI, create a new message.

    ![A message window in Outlook on the new Mac UI with the subject set on compose.](../images/outlook-mac-autolaunch.png)

1. In Outlook on Windows, create a new message.

    ![A message window in Outlook on Windows with the subject set on compose.](../images/outlook-win-autolaunch.png)

## Troubleshooting guide

As you develop your event-based add-in, you may need to troubleshoot issues, such as your add-in not loading or the event not occurring. If you run into development issues, refer to the next four sections for troubleshooting guidance.

### Review event-based activation prerequisites

- Verify that the add-in is installed on a supported Outlook client. Event-based activation isn't supported in Outlook on iOS or Android at this time.
- Verify that your Outlook client supports the minimum requirement set needed to handle the event. Event-based activation was introduced in [requirement set 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), with additional events now supported in subsequent requirements sets. For more information, see [Supported events](#supported-events) and [Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients). If you're developing an add-in that uses the [Smart Alerts](smart-alerts-onmessagesend-walkthrough.md) feature, see the [Supported clients and platform section](smart-alerts-onmessagesend-walkthrough.md#supported-clients-and-platforms).
- Review the expected behavior and limitations of the [event-based activation](#event-based-activation-behavior-and-limitations) and [Smart Alerts](smart-alerts-onmessagesend-walkthrough.md#smart-alerts-feature-behavior-and-scenarios) features.

### Check manifest and JavaScript requirements

- Ensure that the following conditions are met in your add-in's manifest.
  - Verify that your add-in's source file location URL is publicly available and isn't blocked by a firewall. This URL is specified in your manifest's [SourceLocation element](/javascript/api/manifest/sourcelocation).
  - Verify that the **\<Runtimes\>** element correctly references the HTML or JavaScript file containing the event handlers. Outlook on Windows uses the JavaScript file during runtime, while Outlook on the web and on new Mac UI use the HTML file. For an example of how this is configured in the manifest, see [Configure the manifest](#configure-the-manifest).
- Verify that your event-handling JavaScript file referenced by the Outlook client on Windows calls `Office.actions.associate`. This ensures that the event handler name specified in the manifest's **\<LaunchEvent\>** element is mapped to its JavaScript counterpart.

  > [!TIP]
  > If your event-based add-in has only one JavaScript file referenced by Outlook on the web, Windows, and Mac, it's recommended to check on which platform the add-in is running to determine when to call `Office.actions.associate`, as shown in the following code.
  >
  > ```js
  > if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  >   Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  >   Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
  > }
  > ```

- The JavaScript code of event-based add-ins that run in Outlook on Windows only supports [ECMAScript 2016](https://262.ecma-international.org/7.0/) and earlier specifications. Some examples of programming syntax to avoid are as follows.
  - Avoid using `async` and `await` statements in your code. Including these in your JavaScript code will cause the add-in to time out.
  - Avoid using the [conditional (ternary) operator](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Conditional_Operator) as it will prevent your add-in from loading.
  
  If your add-in has only one JavaScript file referenced by Outlook on the web, Windows, and Mac, you must limit your code to ECMAScript 2016 to ensure that your add-in runs in Outlook on Windows. However, if you have a separate JavaScript file referenced by Outlook on the web and Mac, you can implement a later ECMAScript specification in that file.

### Debug your add-in

- As you make changes to your add-in, be aware that:
  - If you update the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in), then sideload it again. If you're using Outlook on Windows, you must also close and reopen Outlook.
  - If you make changes to files other than the manifest, close and reopen the Outlook desktop client, or refresh the browser tab running Outlook on the web.
  - If you're still unable to see your changes after performing these steps, [clear your Office cache](../testing/clear-cache.md).
- As you test your add-in in Outlook on Windows:
  - Check [Event Viewer](/shows/inside/event-viewer) for any reported add-in errors.
    1. In Event Viewer, select **Windows Logs** > **Application**.
    1. From the **Actions** panel, select **Filter Current Log**.
    1. From the **Logged** dropdown, select your preferred log time frame.
    1. Select the **Error** checkbox.
    1. In the **Event IDs** field, enter **63**.
    1. Select **OK** to apply your filters.

    :::image type="content" source="../images/outlook-event-based-logs.png" alt-text="A sample of Event Viewer's Filter Current Log settings configured to only show Outlook errors with event ID 63 that occurred in the last hour.":::

  - Verify that the **bundle.js** file is downloaded to the following folder in File Explorer. Replace text enclosed in `[]` with your applicable information.
  
    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[Outlook mail account encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]
    ```

    If the **bundle.js** file doesn't appear in the folder, [remove your add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) from Outlook, then sideload it again.
- As you test your add-in in Outlook on Windows or Mac, enable runtime logging to identify possible manifest and add-in installation issues. For guidance on how to use runtime logging, see [Debug your add-in with runtime logging](../testing/runtime-logging.md).
- Set breakpoints in your code to debug your add-in. For platform-specific instructions, see [Debug your event-based Outlook add-in](debug-autolaunch.md).

### Seek additional help

If you still need help after performing the recommended troubleshooting steps, [open a GitHub issue](https://github.com/OfficeDev/office-js/issues/new?assignees=&labels=&template=bug_report.md&title=). Include screenshots, video recordings, or runtime logs to supplement your report.

## Deploy to users

Event-based add-ins are restricted to admin-managed deployments only, even if they're acquired from AppSource. If users acquire the add-in from AppSource or the in-app Office Store, they won't be able to activate the event-based function of the add-in. To learn more about listing your event-based add-in in AppSource, see [AppSource listing options for your event-based Outlook add-in](autolaunch-store-options.md).

Admin deployments are done by uploading the manifest to the Microsoft 365 admin center. In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**. On the **Integrated apps** page, choose the **Upload custom apps** action.

![The Integrated apps page on the Microsoft 365 admin center with the Upload custom apps action highlighted.](../images/outlook-deploy-event-based-add-ins.png)

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

### Deploy manifest updates

Because event-based add-ins are deployed by admins, any change you make to the manifest requires admin consent through the Microsoft 365 admin center. Until the admin accepts your changes, users in their organization are blocked from using the add-in. To learn more about the admin consent process, see [Admin consent for installing event-based add-ins](autolaunch-store-options.md#admin-consent-for-installing-event-based-add-ins).

## Event-based activation behavior and limitations

Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible. After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, your associated event handler must call the `event.completed` method. (Note that code included after the `event.completed` statement isn't guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset. The add-in ends after it times out, or the user closes the compose window or sends the item.

If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order. Currently, only five event-based add-ins can be actively running.

In all supported Outlook clients, the user must remain on the current mail item where the add-in was activated for it to complete running. Navigating away from the current item (for example, switching to another compose window or tab) terminates the add-in operation. The add-in also ceases operation when the user sends the message or appointment they're composing.

When developing an event-based add-in to run in the Outlook on Windows client, be mindful of the following:

- Imports aren't supported in the JavaScript file where you implement the handling for event-based activation.
- Add-ins don't run code included in `Office.onReady()` and `Office.initialize`. We recommend adding any startup logic, such as checking the user's Outlook version, to your event handlers instead.

Some Office.js APIs that change or alter the UI aren't allowed from event-based add-ins. The following are the blocked APIs.

- Under `Office.context.auth`:
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth) is supported in all Outlook versions that support event-based activation and single sign-on (SSO), while [Office.auth](/javascript/api/office/office.auth) is only supported in certain Outlook builds. For more information, see [Enable single sign-on (SSO) in Outlook add-ins that use event-based activation](use-sso-in-event-based-activation.md).
- Under `Office.context.mailbox`:
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Under `Office.context.mailbox.item`:
  - `close`
- Under `Office.context.ui`:
  - `displayDialogAsync`
  - `messageParent`

### Requesting external data

You can request external data by using an API like [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using [XMLHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

Be aware that you must use additional security measures when using XMLHttpRequest objects, requiring [Same Origin Policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) and simple [CORS (Cross-Origin Resource Sharing)](https://developer.mozilla.org/docs/Web/HTTP/CORS).

A [simple CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS#simple_requests) implementation:

- Can't use cookies.
- Only supports simple methods, such as `GET`, `HEAD`, and `POST`.
- Accepts simple headers with field names `Accept`, `Accept-Language`, or `Content-Language`.
- Can use the `Content-Type`, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.
- Can't have event listeners registered on the object returned by `XMLHttpRequest.upload`.
- Can't use `ReadableStream` objects in requests.

> [!NOTE]
> Full CORS support is available in Outlook on the web, Mac, and Windows (starting in version 2201, build 16.0.14813.10000).

## See also

- [Outlook add-in manifests](manifests.md)
- [How to debug event-based add-ins](debug-autolaunch.md)
- [AppSource listing options for your event-based Outlook add-in](autolaunch-store-options.md)
- [Smart Alerts and OnMessageSend walkthrough](smart-alerts-onmessagesend-walkthrough.md)
- [Automatically update your signature when switching between mail accounts (preview)](onmessagefromchanged-onappointmentfromchanged-events.md)
- [Manage the sensitivity label of your message or appointment in compose mode (preview)](sensitivity-label.md)
- Office Add-ins code samples:
  - [Use Outlook event-based activation to encrypt attachments, process meeting request attendees and react to appointment date/time changes](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Use Outlook event-based activation to tag external recipients](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
  - [Use Outlook Smart Alerts](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
