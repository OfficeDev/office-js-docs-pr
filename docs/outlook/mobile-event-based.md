---
title: Implement event-based activation in Outlook mobile add-ins
description: Learn how to develop an Outlook mobile add-in that implements event-based activation.
ms.date: 08/01/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement event-based activation in Outlook mobile add-ins

With the [event-based activation](../develop/event-based-activation.md) feature, develop an add-in to automatically activate and complete operations when certain events occur in Outlook on Android or on iOS, such as composing a new message.

The following sections walk you through how to develop an Outlook mobile add-in that automatically adds a signature to new messages being composed. This highlights a sample scenario of how you can implement event-based activation in your mobile add-in. Significantly enhance the mobile user experience by exploring other scenarios and supported events in your add-in today.

To learn how to implement an event-based add-in for Outlook on the web, on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic), and on Mac, see [Activate add-ins with events](../develop/event-based-activation.md).

> [!NOTE]
> Outlook on Android and on iOS only support up to Mailbox requirement set 1.5. However, to support the event-based activation feature, some APIs from later requirement sets have been enabled on mobile clients. For more information on this exception, see [Additional supported APIs](#additional-supported-apis).

## Supported events and clients

| Event canonical name and add-in only manifest name | Unified app manifest for Microsoft 365 name | Description | Supported clients |
| ----- | ----- | ----- | ----- |
| `OnNewMessageCompose` | newMessageComposeCreated | Occurs on composing a new message (includes reply, reply all, and forward), but not on editing a draft. | <ul><li>Android (Version 4.2352.0 and later)</li><li>iOS (Version 4.2352.0 and later)</li></ul> |
| `OnMessageRecipientsChanged` | messageRecipientsChanged | Occurs on adding or removing recipients while composing a message.<br><br>Event-specific data object: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true) | <ul><li>Android (Version 4.2425.0 and later)</li><li>iOS (Version 4.2425.0 and later)</li></ul> |
| `OnMessageFromChanged` | messageFromChanged | Occurs on changing the mail account in the **From** field of a message being composed. To learn more, see  [Automatically update your signature when switching between Exchange accounts](onmessagefromchanged-onappointmentfromchanged-events.md). | <ul><li>Android (Version 4.2502.0 and later)</li><li>iOS (Version 4.2502.0 and later)</li></ul>|

## Set up your environment

To run the feature, you must have a supported version of Outlook on Android or on iOS (see [Supported events and clients](#supported-events-and-clients)) and a Microsoft 365 subscription. Then, complete the [Outlook quick start](../quickstarts/outlook-quickstart-yo.md) in which you create an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

The steps for configuring the manifest depend on which type of manifest you selected in the quick start.

# [Unified app manifest for Microsoft 365](#tab/jsonmanifest)

> [!NOTE]
>
> - When developing an event-based add-in to run in Outlook on Android and on iOS, note that the unified manifest for Microsoft 365 can only be used if the add-in handles certain events. To learn which events are supported, see [Supported events and clients](#supported-events-and-clients).
>
> - Add-ins that use the unified manifest for Microsoft 365 aren't directly supported in Outlook on mobile devices. To run this type of add-in on mobile platforms, the add-in must first be published to [Microsoft Marketplace](https://marketplace.microsoft.com/) then deployed in the [Microsoft 365 Admin Center](../publish/publish.md). For more information, see [Support for add-ins with the unified manifest for Microsoft 365](outlook-mobile-addins.md#support-for-add-ins-with-the-unified-manifest-for-microsoft-365).

1. Configure the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array?view=m365-app-prev&preserve-view=true) property just as you would for setting up a function command. For details, see [Configure the runtime for the function command](../develop/create-addin-commands-unified-manifest.md#configure-the-runtime-for-the-function-command).

1. In the [`"extensions.ribbons.contexts"`](/microsoft-365/extensibility/schema/extension-ribbons-array#contexts) array, add `mailRead` as an item. When you're finished, the array should look like the following.

    ```json
    "contexts": [
        "mailRead"
    ],
    ```

1. In the [`"extensions.ribbons.requirements.formFactors"`](/microsoft-365/extensibility/schema/requirements-extension-element#formfactors) array, add `"mobile"` as an item. When you're finished, the array should look like the following.

    ```json
    "formFactors": [
        "mobile",
        <!-- Typically there will be other form factors listed. -->
    ]
    ```

1. Add the following [`"autoRunEvents"`](/microsoft-365/extensibility/schema/element-extensions#autorunevents) array as a property of the object in the [`"extensions"`](/microsoft-365/extensibility/schema/root#extensions) array.

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. Add an object like the following to the `"autoRunEvents"` array. Note the following about this code:

   - The `"events"` property maps handlers to events.
   - The `"events.type"` must be one of the types listed at [Supported events and clients](#supported-events-and-clients).
   - The value of the `"events.actionId"` is the name of a function that you create in [Implement the event handler](#implement-the-event-handler).
   - You can have more than one object in the `"events"` array.

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.5"
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
          ]
      }
    ```

# [Add-in only manifest](#tab/xmlmanifest)

To enable an event-based add-in on Outlook mobile, you must configure the following elements in the `VersionOverridesV1_1` node of the manifest.

- In the [Runtimes](/javascript/api/manifest/runtimes) element, specify the HTML file that references the event-handling JavaScript file.
- Add the [MobileFormFactor](/javascript/api/manifest/mobileformfactor) element to make your add-in available in Outlook mobile.
- Set the `xsi:type` of the [ExtensionPoint](/javascript/api/manifest/extensionpoint) element to [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent). This enables the event-based activation feature in your Outlook mobile add-in.
- In the [LaunchEvent](/javascript/api/manifest/launchevent) element, set the `Type` to `OnNewMessageCompose` and specify the JavaScript function name of the event handler in the `FunctionName` attribute.

1. In your code editor, open the quick start project you created.
1. Open the **manifest.xml** file located at the root of your project.
1. Select the entire `<VersionOverrides>` node (including the open and close tags) and replace it with the following XML.

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
        <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
            <Requirements>
                <bt:Sets DefaultMinVersion="1.5">
                    <bt:Set Name="Mailbox"/>
                </bt:Sets>
            </Requirements>
            <Hosts>
                <Host xsi:type="MailHost">
                    <!-- The HTML file that references or contains the JavaScript event handlers.
                        This is used by Outlook on mobile devices. -->
                    <Runtimes>
                        <Runtime resid="WebViewRuntime.Url">
                        </Runtime>
                    </Runtimes>
                    <!-- Defines the add-in for Outlook on Windows (new and classic), on Mac, and on the web. -->
                    <DesktopFormFactor>
                        <FunctionFile resid="Commands.Url"/>
                        <ExtensionPoint xsi:type="MessageReadCommandSurface">
                            <OfficeTab id="TabDefault">
                                <Group id="msgReadGroup">
                                    <Label resid="GroupLabel"/>
                                    <Control xsi:type="Button" id="msgReadOpenPaneButton">
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
                    </DesktopFormFactor>
                    <!-- Defines the add-in for Outlook mobile. -->
                    <MobileFormFactor>
                        <!-- Configures event-based activation. -->
                        <ExtensionPoint xsi:type="LaunchEvent">
                            <LaunchEvents>
                                <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
                            </LaunchEvents>
                            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
                            <SourceLocation resid="WebViewRuntime.Url"/>
                        </ExtensionPoint>
                    </MobileFormFactor>
                </Host>
            </Hosts>
            <!-- This manifest uses a fictitious web server, contoso.com, to host the add-in's files.
                 Replace these instances with the information of the web server that hosts your add-in's files. -->
            <Resources>
                <bt:Images>
                    <bt:Image id="Icon.16x16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
                    <bt:Image id="Icon.32x32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
                    <bt:Image id="Icon.80x80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
                </bt:Images>
                <bt:Urls>
                    <bt:Url id="Commands.Url" DefaultValue="https://contoso.com/commands.html"/>
                    <bt:Url id="Taskpane.Url" DefaultValue="https://contoso.com/taskpane.html"/>
                    <bt:Url id="WebViewRuntime.Url" DefaultValue="https://contoso.com/commands.html"/>
                </bt:Urls>
                <bt:ShortStrings>
                    <bt:String id="GroupLabel" DefaultValue="Event-based activation on mobile"/>
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

---

> [!TIP]
> To learn more about manifests for Outlook add-ins, see [Office Add-ins manifest](../develop/add-in-manifests.md) and [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md).

## Implement the event handler

To enable your add-in to complete tasks when the `OnNewMessageCompose` event occurs, you must implement a JavaScript event handler. In this section, you'll create the `onNewMessageComposeHandler` function that adds a signature to a new message being composed, then shows a message to notify that the signature was added.

1. From the same quick start project, navigate to the **./src** directory, then create a new folder named **launchevent**.
1. In the **./src/launchevent** folder, create a new file named **launchevent.js**.
1. Open the **launchevent.js** file you created and add the following JavaScript code.

    ```javascript
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */
    
    // Add start-up logic code here, if any.
    Office.onReady();
    
    function onNewMessageComposeHandler(event) {
        const item = Office.context.mailbox.item;
        const signatureIcon = "iVBORw0KGgoAAAANSUhEUgAAACcAAAAnCAMAAAC7faEHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAzUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKMFRskAAAAQdFJOUwAQIDBAUGBwgI+fr7/P3+8jGoKKAAAACXBIWXMAAA7DAAAOwwHHb6hkAAABT0lEQVQ4T7XT2ZalIAwF0DAJhMH+/6+tJOQqot6X6joPiouNBo3w9/Hd6+hrYnUt6vhLcjEAJevVW0zJxABSlcunhERpjY+UKoNN5+ZgDGu2onNz0OngjP2FM1VdyBW1LtvGeYrBLs7U5I1PTXZt+zifcS3Icw2GcS3vxRY3Vn/iqx31hUyTnV515kdTfbaNhZLI30AceqDiIo4tyKEmJpKdP5M4um+nUwfDWxAXdzqMNKQ14jLdL5ntXzxcRF440mhS6yu882Kxa30RZcUIjTCJg7lscsR4VsMjfX9Q0Vuv/Wd3YosD1J4LuSRtaL7bzXGN1wx2cytUdncDuhA3fu6HPTiCvpQUIjZ3sCcHVbvLtbNTHlysx2w9/s27m9gEb+7CTri6hR1wcTf2gVf3wBRe3CMbcHYvTODkXhnD0+178K/pZ9+n/C1ru/2HAPwAo7YM1X4+tLMAAAAASUVORK5CYII=";
    
        // Get the sender's account information.
        item.from.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.log(result.error.message);
                event.completed();
                return;
            }
    
            // Create a signature based on the sender's information.
            const name = result.value.displayName;
            const options = { asyncContext: name, isInline: true };
            item.addFileAttachmentFromBase64Async(signatureIcon, "signatureIcon.png", options, (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.log(result.error.message);
                    event.completed();
                    return;
                }
    
                // Add the created signature to the message.
                const signature = "<img src='cid:signatureIcon.png'>" + result.asyncContext;
                item.body.setSignatureAsync(signature, { coercionType: Office.CoercionType.Html }, (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.log(result.error.message);
                        event.completed();
                        return;
                    }
    
                    // Show a notification when the signature is added to the message.
                    const notification = {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: "Company signature added.",
                        icon: "none",
                        persistent: false                        
                    };
                    item.notificationMessages.addAsync("signature_notification", notification, (result) => {
                        if (result.status === Office.AsyncResultStatus.Failed) {
                            console.log(result.error.message);
                            event.completed();
                            return;
                        }
    
                        event.completed();
                    });
                });
            });
        });
    }
    ```

1. Save your changes.

## Add a reference to the event-handling JavaScript file

Ensure that the **./src/commands/commands.html** file has a reference to the JavaScript file that contains your event handler.

1. Navigate to the **./src/commands** folder, then open **commands.html**.
1. Immediately before the closing **head** tag (`</head>`), add a script entry for the JavaScript file that contains the event handler.

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. Save you changes.

## Test and validate your add-in

1. Follow the guidance to [test and validate your add-in](testing-and-tips.md).
1. [Sideload](sideload-outlook-add-ins-for-testing.md) your add-in in Outlook on Windows (new or classic), on Mac, or on the web.
1. Open Outlook on Android or on iOS. If you have Outlook already open on your device, restart it.
1. Create a new message. The event-based add-in adds the signature to the message. If you have a signature saved on your mobile device, it will briefly appear in the message you create, but will be immediately replaced by the signature added by the add-in.

    :::image type="content" source="../images/outlook-mobile-signature.png" alt-text="A sample signature added to a message being composed in Outlook mobile.":::

## Behavior and limitations

As you develop an event-based add-in for Outlook mobile, be mindful of the following feature behaviors and limitations.

- Because event-based add-ins are expected to be short-running and lightweight, an add-in is allowed to run for a maximum of 60 seconds from the time it activates. To signal that your add-in has completed processing an event, your event handler must call the [event.completed](/javascript/api/outlook/office.mailboxevent#outlook-office-mailboxevent-completed-member(1)) method. The add-in operation also ends when the user closes the compose window or sends the message.
- Only one add-in can run at a time. If multiple event-based add-ins are installed on a user's account, they will run sequentially.
- If you tap and hold the Outlook icon on your mobile device, then select **New mail** to create a new message, an event-based add-in that handles the `OnNewMessageCompose` event may take a few seconds to initialize and complete processing the event.  
- When using an event-based add-in that handles the `OnNewMessageCompose` event, if there are no changes made to a new message being composed, a draft won't be saved. This applies even if the add-in adds a signature using the [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body#outlook-office-body-setsignatureasync-member(1)) method.
- In an event-based add-in that manages signatures when the `OnNewMessageCompose` event occurs, if you select **Reply** from the bottom of a message, the add-in activates and adds the signature to the message. However, the signature won't be visible in the current view. To view your message with the added signature, expand the compose window to full screen.
- To enhance your add-in's functionality, you can use supported APIs from later requirement sets in compose mode. For more information, see [Additional supported APIs](#additional-supported-apis).

## Additional supported APIs

Although Outlook mobile supports APIs up to [Mailbox requirement set 1.5](/javascript/api/requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5), to further extend the capability of your event-based add-in in Outlook mobile, additional APIs from later requirement sets are now supported in compose mode.

- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-addfileattachmentfrombase64async-member(1))
- [Office.context.mailbox.item.disableClientSignatureAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-disableclientsignatureasync-member(1))
- [Office.context.mailbox.item.from.getAsync](/javascript/api/outlook/office.from#outlook-office-from-getasync-member(1))
- [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getcomposetypeasync-member(1))
- [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body#outlook-office-body-setsignatureasync-member(1))
- [Office.context.mailbox.item.sessionData](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-sessiondata-member)

To learn more about APIs that are supported in Outlook on mobile devices, see [Outlook JavaScript APIs supported in Outlook on mobile devices](outlook-mobile-apis.md).

## Deploy to users

Event-based add-ins must be deployed by an organization's administrator. For guidance on how to deploy your add-in via the Microsoft 365 admin center, see the "Deploy your add-in" section of [Activate add-ins with events](../develop/event-based-activation.md#deploy-your-add-in).

## See also

- [Activate add-ins with events](../develop/event-based-activation.md)
- [Add-ins for Outlook on mobile devices](outlook-mobile-addins.md)
- [Add mobile support to an Outlook add-in](add-mobile-support.md)
- [Outlook JavaScript APIs supported in Outlook on mobile devices](outlook-mobile-apis.md)
