---
title: Automatically update your signature when switching between Exchange accounts (preview)
description: Learn how to automatically update your signature when switching between Exchange accounts through the OnMessageFromChanged event in your event-based activation Outlook add-in.
ms.date: 12/20/2022
ms.localizationpriority: medium
---

# Automatically update your signature when switching between Exchange accounts (preview)

Applying the correct signature to messages when using multiple Exchange accounts is now made easier with the addition of the `OnMessageFromChanged` event to the [event-based activation](autolaunch.md) feature. The `OnMessageFromChanged` event allows your add-in to detect when the account in the **From** field of a message being composed is changed. This event further extends the capabilities of signature add-ins and allows them to:

- Provide users with the convenience to apply custom signatures for each of their accounts.
- Enable mailbox delegates to more accurately and efficiently manage outgoing messages from multiple mailboxes.
- Ensure that users' messages meet their organization's communication and marketing policies.

The following sections walk you through how to develop an event-based add-in that handles the `OnMessageFromChanged` event to automatically update a message's signature when the mail account in the **From** field is changed.

> [!IMPORTANT]
> Features in preview shouldn't be used in production add-ins. We invite you to test this feature in test or development environments and welcome feedback on your experience through GitHub (see the **Feedback** section at the end of this page).

> [!IMPORTANT]
> The `OnMessageFromChanged` event isn't yet supported for the [Teams manifest for Office Add-ins (preview)](../develop/json-manifest-overview.md).

## Supported clients and platforms

The following table lists client-server combinations that support the `OnMessageFromChanged` event.

|Client|Exchange Online|
|-----|-----|
|**Windows**<br>Version 2212 (Build 15919.10000) or later|Yes|
|**Mac**|Not applicable|
|**Web browser (modern UI)**|Not applicable|
|**iOS**|Not applicable|
|**Android**|Not applicable|

## Prerequisites

To test the `OnMessageFromChanged` event while it's in preview, install Outlook on Windows, starting with Version 2212 (Build 15919.10000). Once installed, join the [Office Insider program](https://insider.office.com/join/windows) and select the **Beta Channel** option to access Office beta builds.

To test the walkthrough, you must also have at least two Exchange accounts.

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator), which creates an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

To enable the add-in to activate when the `OnMessageFromChanged` event occurs, the [Runtimes](/javascript/api/manifest/runtimes) element and [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) extension point must be configured in the `VersionOverridesV1_1` node of the manifest.

In addition to the `OnMessageFromChanged` event, the `OnNewMessageCompose` event is also configured in the manifest, so that a signature can be added to a message being composed if a default Outlook signature isn't already configured on the current account.

1. In your code editor, open the quick start project.

1. Open the **manifest.xml** file located at the root of your project.

1. Select the entire **\<VersionOverrides\>** node (including open and close tags), replace it with the following XML, then save your changes.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
       <Requirements>
         <bt:Sets DefaultMinVersion="1.12">
           <bt:Set Name="Mailbox"/>
         </bt:Sets>
       </Requirements>
       <Hosts>
         <Host xsi:type="MailHost">
           <Runtimes>
             <!-- HTML file that references or contains inline JavaScript event handlers.
                  This is used by event-based activation add-ins in Outlook on the web and Outlook on the new Mac UI. -->
             <Runtime resid="WebViewRuntime.Url">
               <!-- JavaScript file that contains the event handlers.
                    This is used by event-based activation add-ins in Outlook on Windows. -->
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

> [!TIP]
>
> - To learn about runtimes in add-ins, see [Runtimes in Office Add-ins](../testing/runtimes.md).
> - To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).

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
        const item = Office.context.mailbox.item;
    
        // Check if a default Outlook signature is already configured.
        item.isClientSignatureEnabledAsync({ asyncContext: event }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.log(result.error.message);
                return;
            }
    
            // Add a signature if there's no default Outlook signature configured.
            if (result.value === false) {
                item.body.setSignatureAsync(
                    "<i>This is a sample signature.</i>",
                    { asyncContext: result.asyncContext, coercionType: Office.CoercionType.Html },
                    addSignatureCallback
                );
            }
        });
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
                item.body.setSignatureAsync(
                    signature,
                    { asyncContext: result.asyncContext.event, coercionType: Office.CoercionType.Html },
                    addSignatureCallback
                );
            });
        });
    }
    
    // Callback function to add a signature to the mail item.
    function addSignatureCallback(result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.log(result.error.message);
            return;
        }
    
        console.log("Successfully added signature.");
        result.asyncContext.completed();
    }
    
    // IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to 
    // map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
    if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
        Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
        Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
    }
    ```

> [!IMPORTANT]
> Windows: At present, imports aren't supported in the JavaScript file where you implement the handling for event-based activation.

> [!TIP]
> Event-based add-ins running in Outlook on Windows don't run code included in the `Office.onReady()` and `Office.initialize` functions. We recommend adding your add-in startup logic, such as checking the user's Outlook version, to your event handlers instead.

## Update the commands HTML file

1. From the **./src/commands** folder, open **commands.html**.

1. Replace the existing **script** tag with the following reference to the beta library on the content delivery network (CDN).

   ```html
   <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"></script>
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

    > [!NOTE]
    > If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.

1. In Outlook on Windows, create a new message. If you don't have a default Outlook signature configured, the add-in adds one to the newly created message.

   :::image type="content" source="../images/OnMessageFromChanged_create_signature.png" alt-text="A sample signature added to a newly composed message when a default Outlook signature isn't configured on the account.":::

1. Enable the **From** field, if applicable. For guidance on how to enable it, see the "Why is the From button missing?" section of [Change the account used to send email messages](https://support.microsoft.com/office/2bdd8d4f-e30f-4ec8-88a0-406ce7b23cc5).

1. Select **From**, then choose a different Exchange account. Alternatively, manually enter the Exchange email address by selecting **From** > **Other Email Address**. An updated signature is added to the message, replacing the previous one.

   :::image type="content" source="../images/OnMessageFromChanged_update_signature.png" alt-text="A sample of an updated signature with a logo when the account in the From field is changed.":::

## Troubleshoot your add-in

For guidance on how to troubleshoot your event-based activation add-in, see the "Troubleshooting guide" section of [Configure your Outlook add-in for event-based activation](autolaunch.md#troubleshooting-guide).

## Deploy to users

Similar to other event-based add-ins, add-ins that use the `OnMessageFromChanged` event must be deployed by an organization's administrator. For guidance on how to deploy your add-in via the Microsoft 365 admin center, see the "Deploy to users" section of [Configure your Outlook add-in for event-based activation](autolaunch.md#deploy-to-users).

## Event behavior and limitations

Because the `OnMessageFromChanged` event is supported through the event-based activation feature, the same behavior and limitations apply to add-ins that activate as a result of this event. For a detailed description, see [Event-based activation behavior and limitations](autolaunch.md#event-based-activation-behavior-and-limitations).

In addition to these characteristics, the following aspects also apply when an add-in activates on the `OnMessageFromChanged` event.

- The `OnMessageFromChanged` event is only supported in message compose mode.
- The `OnMessageFromChanged` event only supports Exchange accounts. The Exchange account can be selected from the **From** field dropdown or manually entered in the field. Non-Exchange accounts aren't supported. If a user switches to a non-supported account in the **From** field, the Outlook client automatically clears out the signature set by the previously selected account.
- Delegate and shared mailbox scenarios are supported.
- When switching to an Exchange account in the **From** field, the add-ins for the previously selected account, if any, are terminated, and the add-ins associated with the newly selected account are loaded before the `OnMessageFromChanged` event is initiated.
- Email account aliases are supported. When an alias for the current account is selected in the **From** field, the `OnMessageFromChanged` event occurs without reloading the account's add-ins.
- When the **From** field dropdown is opened by mistake or the same account that appears in the **From** field is reselected, the `OnMessageFromChanged` event occurs, but the account's add-ins aren't terminated or reloaded.

## See also

- [Configure your Outlook add-in for event-based activation](autolaunch.md)
- [AppSource listing options for your event-based Outlook add-in](autolaunch-store-options.md)
