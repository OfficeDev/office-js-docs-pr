---
title: Automatically set the subject of a new message or appointment
description: Learn how to implement an event-based add-in that automatically sets the subject of a new message or appointment.
ms.date: 01/13/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Automatically set the subject of a new message or appointment

Need to add a required disclaimer to all your messages? With an event-based add-in, content is automatically added to new messages or appointments. Your users can focus on writing, instead of compliance.

The following sections teach you how to develop an add-in that handles the `OnNewMessageCompose` and `OnNewAppointmentOrganizer` events. By the end of this walkthrough, you'll have an add-in that automatically sets the subject of new messages and appointments being created.

> [!NOTE]
>
> - The `OnNewMessageCompose` and `OnNewAppointmentOrganizer` events were introduced in [requirement set 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10). To verify that your Outlook client supports these events, see [Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).
> - The `OnNewMessageCompose` event is now supported in Outlook on mobile devices. To learn how to implement this event in your Outlook mobile add-in, see [Implement event-based activation in Outlook mobile add-ins](mobile-event-based.md).

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart-yo.md) which creates an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

To configure the manifest, select the tab for the type of manifest you're using.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

[!INCLUDE [outlook-unified-manifest-platforms](../includes/outlook-unified-manifest-platforms.md)]

1. Open the **manifest.json** file.

1. Navigate to the [`"authorization.permissions.resourceSpecific"`](/microsoft-365/extensibility/schema/root-authorization-permissions#resourcespecific) array. In the array object, replace the value of the `"name"` property with `"MailboxItem.ReadWrite.User"`. This is needed by the add-in to be able to set the subject of the mail item.

    ```json
    ...
    "authorization": {
        "permissions": {
            "resourceSpecific": [
                {
                    "name": "MailboxItem.ReadWrite.User",
                    "type": "Delegated"
                }
            ]
        }
    },
    ...
    ```

1. Add the following object to the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array?view=m365-app-prev&preserve-view=true) array. Note the following about this markup:

   - The `"minVersion"` of the Mailbox requirement set is configured to `"1.10"` as this is the lowest version of the requirement set that supports the `OnNewMessageCompose` and `OnNewAppointmentOrganizer` events.
   - The `"id"` of the runtime is set to the descriptive name `"autorun_runtime"`.
   - The `"code"` property has a child `"page"` property that is set to an HTML file and a child `"script"` property that is set to a JavaScript file. You'll create or edit these files in later steps. Office uses one of these values depending on the platform.
       - Office on Windows executes the event handlers in a JavaScript-only runtime, which loads a JavaScript file directly.
       - Office on Mac and on the web, and [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) execute the handlers in a browser runtime, which loads an HTML file. That file, in turn, contains a `<script>` tag that loads the JavaScript file.

     For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md).
   - The `"lifetime"` property is set to `"short"`, which means that the runtime starts up when one of the events is triggered and shuts down when the handler completes. (In certain rare cases, the runtime shuts down before the handler completes. See [Runtimes in Office Add-ins](../testing/runtimes.md).)
   - There are two types of [`"actions"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item) that can run in the runtime. You'll create functions to correspond to these actions in a later step.

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
                "type": "executeFunction"
            },
            {
                "id": "onNewAppointmentComposeHandler",
                "type": "executeFunction"
            }
        ]
    }
    ```

1. Add the following [`"autoRunEvents"`](/microsoft-365/extensibility/schema/element-extensions#autorunevents) array as a property of the object in the [`"extensions"`](/microsoft-365/extensibility/schema/root#extensions) array.

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. Add the following object to the `"autoRunEvents"` array. The `"events"` property maps handlers to events as described in the table earlier in this article. The handler names must match those used in the `"id"` properties of the objects in the `"actions"` array in an earlier step.

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

# [Add-in only manifest](#tab/xmlmanifest)

To enable event-based activation of your add-in, you must configure the [Runtimes](/javascript/api/manifest/runtimes) element and [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) extension point in the `VersionOverridesV1_1` node of the manifest.

In event-based add-ins, classic Outlook on Windows uses a JavaScript file, while Outlook on the web and on the new Mac UI, and [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) use an HTML file that can reference the same JavaScript file. You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client. As such, to configure event handling, provide the location of the HTML in the `<Runtime>` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.

1. In your code editor, open the quick start project.

1. Open the **manifest.xml** file located at the root of your project.

1. Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.

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
               This is used by Outlook on the web and on the new Mac UI, and new Outlook on Windows. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by classic Outlook on Windows. -->
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
        <!-- Entry needed for classic Outlook on Windows. -->
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

---

> [!TIP]
>
> - To learn about runtimes in add-ins, see [Runtimes in Office Add-ins](../testing/runtimes.md).
> - To learn more about manifests for Outlook add-ins, see [Office Add-in manifests](../develop/add-in-manifests.md).

## Implement event handling

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

          // Call event.completed() to signal to the Outlook client that the add-in has completed processing the event.
          asyncResult.asyncContext.completed();
        });
    }

    // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
    ```

1. Save your changes.

> [!NOTE]
>
> - There are some limitations you must be aware of when developing an event-based add-in for classic Outlook on Windows. To learn more, see [Event-based activation behavior and limitations](../develop/event-based-activation.md#behavior-and-limitations).
> - To ensure your add-in runs as expected when an event occurs, call `Office.actions.associate` in the JavaScript file where your handlers are implemented. This maps the event handler name specified in the manifest to its JavaScript counterpart. The location of the handler name in the manifest differs depending on the type of manifest your add-in uses.
>   - **Unified manifest for Microsoft 365**: The value specified in the [`"actionId"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events#actionid) property of the applicable [`"autoRunEvents.events"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events) object.
>   - **Add-in only manifest**: The function name specified in the applicable [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) element.
> - In classic Outlook on Windows, when composing a message initiated by a [return email link](https://support.microsoft.com/office/86cea017-8f4e-4f20-85aa-0683779ccb0a) (`mailto` link), retrieving the recipients from the To, Cc, or Bcc field in the `OnNewMessageCompose` event handler may return an empty array. This happens if Outlook hasn't completed resolving the email addresses of the recipients by the time the `OnNewMessageCompose` event occurs. To work around this, check for recipients in the `OnMessageRecipientsChanged` event handler instead. For more information about supported events, see [Outlook events](../develop/event-based-activation.md#outlook-events).

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

    [!INCLUDE [outlook-manual-sideloading](../includes/outlook-manual-sideloading.md)]

1. In Outlook on the web or in new Outlook on Windows, create a new message.

    :::image type="content" source="../images/outlook-web-autolaunch-1.png" alt-text="A message window in Outlook on the web with the subject set on compose.":::

1. In Outlook on the new Mac UI, create a new message.

    :::image type="content" source="../images/outlook-mac-autolaunch.png" alt-text="A message window in Outlook on the new Mac UI with the subject set on compose.":::

1. In classic Outlook on Windows, create a new message.

    :::image type="content" source="../images/outlook-win-autolaunch.png" alt-text="A message window in classic Outlook on Windows with the subject set on compose.":::

1. [!include[Instructions to stop web server and uninstall dev add-in](../includes/stop-uninstall-outlook-dev-add-in.md)]

## Next steps

To learn more about event-based activation and other events that you can implement in your add-in, see [Activate add-ins with events](../develop/event-based-activation.md)

## See also

- [Activate add-ins with events](../develop/event-based-activation.md)
- [Office Add-ins code sample: Set your signature using Outlook event-based activation](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
