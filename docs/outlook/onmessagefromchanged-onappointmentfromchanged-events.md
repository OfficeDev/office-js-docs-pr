---
title: Automatically update your signature when switching between Exchange accounts
description: Learn how to automatically update your signature when switching between Exchange accounts through the OnMessageFromChanged and OnAppointmentFromChanged events in your event-based activation Outlook add-in.
ms.date: 07/17/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Automatically update your signature when switching between Exchange accounts

Applying the correct signature to messages when using multiple Exchange accounts is now made easier with the addition of the `OnMessageFromChanged` and `OnAppointmentFromChanged` events to the [event-based activation](../develop/event-based-activation.md) feature. The `OnMessageFromChanged` event occurs when the account in the **From** field of a message being composed is changed, while the `OnAppointmentFromChanged` event occurs when the organizer of a meeting being composed is changed. These events further extend the capabilities of signature add-ins and allow them to:

- Provide users with the convenience to apply custom signatures for each of their accounts.
- Enable mailbox delegates to more accurately and efficiently manage outgoing messages and meeting requests from multiple mailboxes.
- Ensure that users' messages and appointments meet their organization's communication and marketing policies.

The following sections walk you through how to develop an event-based add-in that handles the `OnMessageFromChanged` event to automatically update a message's signature when the mail account in the **From** field is changed.

> [!NOTE]
> The `OnMessageFromChanged` and `OnAppointmentFromChanged` events were introduced in [requirement set 1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13). For information about client support for these events, see [Supported clients and platforms](#supported-clients-and-platforms).

## Supported clients and platforms

The following tables list client-server combinations that support the `OnMessageFromChanged` and `OnAppointmentFromChanged` events. Select the tab for the applicable event.

# [OnMessageFromChanged event](#tab/message)

|Client|Exchange Online|Exchange 2019 on-premises (Cumulative Update 12 or later)|Exchange 2016 on-premises (Cumulative Update 22 or later)|
|-----|-----|-----|-----|
|**Web browser (modern UI)**<br><br>[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)|Supported|Not applicable|Not applicable|
|**Windows (classic)**<br>Version 2304 (Build 16327.20248) or later|Supported|Supported|Supported|
|**Mac**<br>Version 16.77 (23081600) or later|Supported|Not applicable|Not applicable|
|**iOS**<br>Version 4.2502.0|Supported|Not applicable|Not applicable|
|**Android**<br>Version 4.2502.0|Supported|Not applicable|Not applicable|

# [OnAppointmentFromChanged event](#tab/appointment)

|Client|Exchange Online|Exchange 2019 on-premises (Cumulative Update 12 or later)|Exchange 2016 on-premises (Cumulative Update 22 or later)|
|-----|-----|-----|-----|
|**Web browser (modern UI)**<br><br>[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)|Not applicable|Not applicable|Not applicable|
|**Windows (classic)**|Not applicable|Not applicable|Not applicable|
|**Mac**<br>Version 16.77 (23081600) or later|Supported|Not applicable|Not applicable|
|**iOS**|Not applicable|Not applicable|Not applicable|
|**Android**|Not applicable|Not applicable|Not applicable|

---

## Prerequisites

To test the walkthrough, you must have at least two Exchange accounts.

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart-yo.md), which creates an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

> [!NOTE]
> The `OnMessageFromChanged` event isn't currently available with the unified manifest for Microsoft 365 in Outlook on the web and the new Outlook on Windows. To handle these events, implement an add-in only manifest instead. For information about the types of manifests, see [Office Add-ins manifest](../develop/add-in-manifests.md).

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

1. Open the **manifest.json** file.

1. Navigate to the [`"authorization.permissions.resourceSpecific"`](/microsoft-365/extensibility/schema/root-authorization-permissions-resource-specific) array. In the array object, replace the value of the [`"name"`](/microsoft-365/extensibility/schema/root-authorization-permissions-resource-specific#name) property with `"MailboxItem.ReadWrite.User"`. This is needed by the add-in to be able to update the signature of a message.

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

1. Add the following object to the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array) array. Note the following about this markup.

   - The [`"minVersion"`](/microsoft-365/extensibility/schema/requirements-extension-element-capabilities#minversion) of the Mailbox requirement set is configured as `"1.13"` because this is the lowest version of the requirement set that supports the `OnMessageFromChanged` event. For more information, see the "Supported events" table in [Activate add-ins with events](../develop/event-based-activation.md#supported-events).
   - The [`"id"`](/microsoft-365/extensibility/schema/extension-runtimes-array#id) of the runtime is set to a descriptive name, `"autorun_runtime"`.
   - The [`"code"`](/microsoft-365/extensibility/schema/extension-runtime-code) property has a child [`"page"`](/microsoft-365/extensibility/schema/extension-runtime-code#page) property set to an HTML file and a child [`"script"`](/microsoft-365/extensibility/schema/extension-runtime-code#script) property set to a JavaScript file. You'll create or edit these files in later steps. Office uses one of these values depending on the platform.
       - Classic Outlook on Windows executes the event handler in a JavaScript-only runtime, which loads a JavaScript file directly.
       - Outlook on the web and on Mac, and new Outlook on Windows execute the handler in a browser runtime, which loads an HTML file. The HTML file contains a `<script>` tag that then loads the JavaScript file.

     For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md).
   - The [`"lifetime"`](/microsoft-365/extensibility/schema/extension-runtimes-array#lifetime) property is set to `"short"`. This means the runtime starts up when the event occurs and shuts down when the handler completes.
   - There are [`"actions"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item) to run handlers for the `OnMessageFromChanged` and `OnNewMessageCompose` events. You'll create the handlers in a later step.

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.13"
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
                "id": "onMessageFromChangedHandler",
                "type": "executeFunction",
                "displayName": "onMessageFromChangedHandler"
            },
            {
                "id": "onNewMessageComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewMessageComposeHandler"
            }
        ]
    }
    ```

1. Add an [`"autoRunEvents"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array) array as a property of the object in the [`"extensions"`](/microsoft-365/extensibility/schema/root#extensions) array. The `"autoRunEvents"` array contains an object with the following key properties.

    - The [`"events"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events) property assigns handlers to the `OnMessageFromChanged` and `OnNewMessageCompose` events. For information on event names used in the unified manifest, see the "Supported events" table in [Activate add-ins with events](../develop/event-based-activation.md#supported-events).
    - The function name provided in [`"actionId"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events#actionid) must match the `"id"` property of its corresponding object in the `"actions"` array configured earlier.

    ```json
    "autoRunEvents": [
        {
            "requirements": {
                "capabilities": [
                    {
                        "name": "Mailbox",
                        "minVersion": "1.13"
                    }
                ],
                "scopes": [
                    "mail"
                ]
            },
            "events": [
                {
                    "type": "messageFromChanged",
                    "actionId": "onMessageFromChangedHandler"
                },
                {
                    "type": "newMessageComposeCreated",
                    "actionId": "onNewMessageComposeHandler"
                }
            ]
        }
    ]
    ```

# [Add-in only manifest](#tab/xmlmanifest)

To enable the add-in to activate when the `OnMessageFromChanged` event occurs, the [Runtimes](/javascript/api/manifest/runtimes) element and [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) extension point must be configured in the `VersionOverridesV1_1` node of the manifest.

In addition to the `OnMessageFromChanged` event, the `OnNewMessageCompose` event is also configured in the manifest, so that a signature is added to a message being composed if a default Outlook signature isn't already configured on the current account.

1. In your code editor, open the quick start project.

1. Open the **manifest.xml** file located at the root of your project.

1. Select the entire **\<VersionOverrides\>** node (including open and close tags), replace it with the following XML, then save your changes.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
       <Requirements>
         <bt:Sets DefaultMinVersion="1.13">
           <bt:Set Name="Mailbox"/>
         </bt:Sets>
       </Requirements>
       <Hosts>
         <Host xsi:type="MailHost">
           <Runtimes>
             <!-- HTML file that references or contains inline JavaScript event handlers.
                  This is used by event-based activation add-ins in Outlook on the web and on Mac,
                  and in new Outlook on Windows. -->
             <Runtime resid="WebViewRuntime.Url">
               <!-- JavaScript file that contains the event handlers.
                    This is used by event-based activation add-ins in classic Outlook on Windows. -->
               <Override type="javascript" resid="JSRuntime.Url"/>
             </Runtime>
           </Runtimes>
           <DesktopFormFactor>
             <FunctionFile resid="Commands.Url"/>
             <ExtensionPoint xsi:type="MessageComposeCommandSurface">
               <OfficeTab id="TabDefault">
                 <Group id="msgComposeGroup">
                   <Label resid="GroupLabel"/>
                   <Control xsi:type="Button" id="msgComposeOpenPaneButton">
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
                       <SourceLocation resid="Taskpane.Url"/>
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
             <!-- Configures event-based activation. -->
             <ExtensionPoint xsi:type="LaunchEvent">
               <LaunchEvents>
                 <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
                 <LaunchEvent Type="OnMessageFromChanged" FunctionName="onMessageFromChangedHandler"/>
               </LaunchEvents>
               <!-- Identifies the runtime to be used (also referenced by the <Runtime> element). -->
               <SourceLocation resid="WebViewRuntime.Url"/>
             </ExtensionPoint>
           </DesktopFormFactor>
           <!-- Defines the add-in for Outlook mobile. -->
           <MobileFormFactor>
             <!-- Configures event-based activation. -->
             <ExtensionPoint xsi:type="LaunchEvent">
               <LaunchEvents>
                 <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
                 <LaunchEvent Type="OnMessageFromChanged" FunctionName="onMessageFromChangedHandler"/>
               </LaunchEvents>
               <!-- Identifies the runtime to be used (also referenced by the <Runtime> element). -->
               <SourceLocation resid="WebViewRuntime.Url"/>
             </ExtensionPoint>
           </MobileFormFactor>
         </Host>
       </Hosts>
       <Resources>
         <bt:Images>
           <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
           <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
           <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
         </bt:Images>
         <bt:Urls>
           <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
           <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
           <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/launchevent.js"/>
           <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html"/>
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

## Implement the event handlers

Event handlers must be configured for the `OnNewMessageCompose` and `OnMessageFromChanged` events. The `onNewMessageComposeHandler` function adds a signature to a newly created message if a default one isn't already configured on the current account. When the account in the **From** field is changed, the `onMessageFromChangedHandler` function updates the signature based on this newly selected account.

1. From the same quick start project, navigate to the **./src** directory, then create a new folder named **launchevent**.

1. In the **./src/launchevent** folder, create a new file named **launchevent.js**.

1. Open the file **./src/launchevent/launchevent.js** in your code editor and add the following JavaScript code.

    ```javascript
    /*
     * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
     * See LICENSE in the project root for license information.
     */

    // The OnNewMessageCompose event handler that adds a signature to a new message.
    function onNewMessageComposeHandler(event) {
        const platform = Office.context.platform;
        const signature = "<i>This is a sample signature.</i>";

        // On supported platforms, check if a default Outlook signature is already configured.
        if (platform !== Office.PlatformType.Android && platform !== Office.PlatformType.iOS) {
            Office.context.mailbox.item.isClientSignatureEnabledAsync({ asyncContext: { event: event, signature: signature } }, (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.log(result.error.message);
                    return;
                }

                // Add a signature if there's no default Outlook signature configured.
                const signatureEnabled = result.value;
                if (signatureEnabled === false) {
                    const event = result.asyncContext.event;
                    const signature = result.asyncContext.signature;
                    setSignature(signature, event);
                }
            });
        } else {
            setSignature(signature, event);
        }
    }

    // The OnMessageFromChanged event handler that updates the signature when the email address in the From field is changed.
    function onMessageFromChangedHandler(event) {
        const item = Office.context.mailbox.item;
        const signatureIcon =
        "iVBORw0KGgoAAAANSUhEUgAAACcAAAAnCAMAAAC7faEHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAzUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKMFRskAAAAQdFJOUwAQIDBAUGBwgI+fr7/P3+8jGoKKAAAACXBIWXMAAA7DAAAOwwHHb6hkAAABT0lEQVQ4T7XT2ZalIAwF0DAJhMH+/6+tJOQqot6X6joPiouNBo3w9/Hd6+hrYnUt6vhLcjEAJevVW0zJxABSlcunhERpjY+UKoNN5+ZgDGu2onNz0OngjP2FM1VdyBW1LtvGeYrBLs7U5I1PTXZt+zifcS3Icw2GcS3vxRY3Vn/iqx31hUyTnV515kdTfbaNhZLI30AceqDiIo4tyKEmJpKdP5M4um+nUwfDWxAXdzqMNKQ14jLdL5ntXzxcRF440mhS6yu882Kxa30RZcUIjTCJg7lscsR4VsMjfX9Q0Vuv/Wd3YosD1J4LuSRtaL7bzXGN1wx2cytUdncDuhA3fu6HPTiCvpQUIjZ3sCcHVbvLtbNTHlysx2w9/s27m9gEb+7CTri6hR1wcTf2gVf3wBRe3CMbcHYvTODkXhnD0+178K/pZ9+n/C1ru/2HAPwAo7YM1X4+tLMAAAAASUVORK5CYII=";

        // Get the currently selected From account.
        item.from.getAsync({ asyncContext: event }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.log(result.error.message);
                return;
            }

            // Create a signature based on the currently selected From account.
            const name = result.value.displayName;
            const options = { asyncContext: { event: result.asyncContext, name: name }, isInline: true };
            item.addFileAttachmentFromBase64Async(signatureIcon, "signatureIcon.png", options, (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.log(result.error.message);
                    return;
                }
        
                // Add the created signature to the mail item.
                const signature = "<img src='cid:signatureIcon.png'>" + result.asyncContext.name;
                const event = result.asyncContext.event;
                setSignature(signature, event);
            });
        });
    }

    // Sets the custom signature and adds it to the mail item.
    function setSignature(signature, event) {
        Office.context.mailbox.item.body.setSignatureAsync(
            signature,
            { asyncContext: event, coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.log(result.error.message);
                    return;
                }

                console.log("Successfully added signature.");
                const event = result.asyncContext;
                event.completed();
            }
        );
    }

    // IMPORTANT: To ensure your add-in is supported in Outlook, remember to
    // map the event handler name specified in the manifest to its JavaScript counterpart.
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
    ```

> [!IMPORTANT]
>
> - In classic Outlook on Windows, imports aren't supported in the JavaScript file where you implement the handling for event-based activation.
> - In classic Outlook on Windows, when the JavaScript function specified in the manifest to handle an event runs, code in `Office.onReady()` and `Office.initialize` isn't run. We recommend adding any startup logic needed by event handlers, such as checking the user's Outlook version, to the event handlers instead.
> - To ensure your add-in runs as expected when an event occurs, call `Office.actions.associate` in the JavaScript file where your handlers are implemented. This maps the event handler name specified in the manifest to its JavaScript counterpart. The location of the handler name in the manifest differs depending on the type of manifest your add-in uses.
>   - **Unified manifest for Microsoft 365**: The value specified in the [`"actionId"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events#actionid) property of the applicable [`"autoRunEvents.events"`](/microsoft-365/extensibility/schema/extension-auto-run-events-array-events) object.
>   - **Add-in only manifest**: The function name specified in the applicable [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) element.

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

1. In your preferred Outlook client, create a new message. If you don't have a default Outlook signature configured, the add-in adds one to the newly created message. In Outlook on mobile devices, the add-in adds a sample signature even if you have a default signature configured.

   :::image type="content" source="../images/OnMessageFromChanged_create_signature.png" alt-text="A sample signature added to a newly composed message when a default Outlook signature isn't configured on the account.":::

1. Enable the **From** field, if applicable. For guidance on how to enable it, see the "Why is the From button missing?" section of [Change the account used to send email messages](https://support.microsoft.com/office/2bdd8d4f-e30f-4ec8-88a0-406ce7b23cc5).

1. Select **From**, then choose a different Exchange account. Alternatively, manually enter the Exchange email address by selecting **From** > **Other Email Address**. An updated signature is added to the message, replacing the previous one.

   :::image type="content" source="../images/OnMessageFromChanged_update_signature.png" alt-text="A sample of an updated signature with a logo when the account in the From field is changed.":::

1. [!include[Instructions to stop web server and uninstall dev add-in](../includes/stop-uninstall-outlook-dev-add-in.md)]

## Troubleshoot your add-in

For guidance on how to troubleshoot your event-based activation add-in, see [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md).

## Deploy to users

Similar to other event-based add-ins, add-ins that use the `OnMessageFromChanged` and `OnAppointmentFromChanged` events must be deployed by an organization's administrator. For guidance on how to deploy your add-in via the Microsoft 365 admin center, see the "Deploy your add-in" section of [Activate add-ins with events](../develop/event-based-activation.md#deploy-your-add-in).

## Event behavior and limitations

Because the `OnMessageFromChanged` and `OnAppointmentFromChanged` events are supported through the event-based activation feature, the same behavior and limitations apply to add-ins that activate as a result of this event. For a detailed description, see [Event-based activation behavior and limitations](../develop/event-based-activation.md#behavior-and-limitations).

In addition to these characteristics, the following aspects also apply when an add-in activates on these events.

- The `OnMessageFromChanged` event is only supported in message compose mode, while the `OnAppointmentFromChanged` event is only supported in appointment compose mode.
- In Outlook on the web, on Windows (new and classic), and on mobile devices, only the `OnMessageFromChanged` event is supported.
- The `OnMessageFromChanged` and `OnAppointmentFromChanged` events only support Exchange accounts. If a user switches to a non-Exchange account in the **From** or organizer field, the Outlook client automatically clears out the signature set by the previously selected account.
- Depending on your Outlook client, in messages being composed, the Exchange account is selected from the **From** field dropdown list or manually entered in the field. Outlook on mobile devices only supports selecting an account from the **From** field dropdown list. In appointments being composed, the Exchange account is selected from the organizer field dropdown list.
- In Outlook on the web, on Windows (new and classic), and on Mac, the `OnMessageFromChanged` and `OnAppointmentFromChanged` events support delegate and shared mailbox scenarios. These scenarios aren't supported in Outlook on mobile devices.
- The `OnAppointmentFromChanged` event isn't supported in [Microsoft 365 group calendars](https://support.microsoft.com/office/0cf1ad68-1034-4306-b367-d75e9818376a#Outlook=Web). If a user switches from their Exchange account to a Microsoft 365 group calendar account in the organizer field, the Outlook client automatically clears out the signature set by the Exchange account.
- When switching to another Exchange account in the **From** or organizer field, the add-ins for the previously selected account, if any, are terminated, and the add-ins associated with the newly selected account are loaded before the `OnMessageFromChanged` or `OnAppointmentFromChanged` event is initiated.
- In Outlook on the web, on Windows (new and classic), and on Mac, email account aliases are supported. When an alias for the current account is selected in the **From** or organizer field, the `OnMessageFromChanged` or `OnAppointmentFromChanged` event occurs without reloading the account's add-ins. Email account aliases aren't supported in Outlook on mobile devices.
- When the **From** or organizer field dropdown list is opened by mistake or the same account that appears in the **From** or organizer field is reselected, the `OnMessageFromChanged` or `OnAppointmentFromChanged` event occurs, but the account's add-ins aren't terminated or reloaded.

## See also

- [Activate add-ins with events](../develop/event-based-activation.md)
- [AppSource listing options for your event-based add-in](../publish/autolaunch-store-options.md)
