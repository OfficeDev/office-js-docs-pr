---
title: Activate your Outlook add-in on multiple messages
description: Learn how to activate your Outlook add-in when multiple messages are selected.
ms.date: 05/20/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Activate your Outlook add-in on multiple messages

With the item multi-select feature, your Outlook add-in can now activate and perform operations on multiple selected messages in one go. Certain operations, such as uploading messages to your Customer Relationship Management (CRM) system or categorizing numerous items, can now be easily completed with a single click.

The following sections walk you through how to configure your add-in to retrieve the subject line of multiple messages in read mode.

> [!NOTE]
> Support for the item multi-select feature was introduced in [requirement set 1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13), with additional item properties now available in subsequent requirement sets. See [clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) to create an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

1. In your preferred code editor, open the Outlook quick start project you created.

1. Open the **manifest.json** file located at the root of the project.

1. In the "authorization.permissions.resourceSpecific" array, change the value of the "name" property to "Mailbox.ReadWrite.User". It should look like the following when you're done.

    ```json
    "authorization": {
        "permissions": {
            "resourceSpecific": [
                {
                    "name": "Mailbox.ReadWrite.User",
                    "type": "Delegated"
                }
            ]
        }
    },
    ```

1. In first object of the "extensions.runtimes" array, make the following changes.

    1. Change the "requirements.capabilities.minVersion" property to "1.13".
    1. In the same "actions" object, add the "supportsNoItemContext" property and set it to `true`.
    1. In the same "actions" object, add the "multiselect" property and set it to `true`.

    Your code should look like the following after you've made the changes.

    ```json
    "runtimes": [
        {
            "requirements": {
                "capabilities": [
                    {
                        "name": "Mailbox",
                        "minVersion": "1.13"
                    }
                ]
            },
            "id": "TaskPaneRuntime",
            "type": "general",
            "code": {
                "page": "https://localhost:3000/taskpane.html"
            },
            "lifetime": "short",
            "actions": [
                {
                    "id": "TaskPaneRuntimeShow",
                    "type": "openPage",
                    "pinnable": false,
                    "view": "dashboard",
                    "supportsNoItemContext": true,
                    "multiselect": true
                }
            ]
        },
        ...
    ]
    ```

1. Delete the second object of the "extensions.runtimes" array, whose "id" is "CommandsRuntime".

1. In the "extensions.ribbons.tabs.controls" array, delete the second object, whose "id" is "ActionButton".

1. Save your changes.

# [XML Manifest](#tab/xmlmanifest)

To enable your add-in to activate on multiple selected messages, you must add the [SupportsMultiSelect](/javascript/api/manifest/action#supportsmultiselect) child element to the **\<Action\>** element and set its value to `true`. As item multi-select only supports messages at this time, the **\<ExtensionPoint\>** element's `xsi:type` attribute value must be set to `MessageReadCommandSurface` or `MessageComposeCommandSurface`.

1. In your preferred code editor, open the Outlook quick start project you created.

1. Open the **manifest.xml** file located at the root of the project.

1. Assign the **\<Permissions\>** element the `ReadWriteMailbox` value.

    ```xml
    <Permissions>ReadWriteMailbox</Permissions>
    ```

1. Select the entire **\<VersionOverrides\>** node and replace it with the following XML.

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
                    <DesktopFormFactor>
                        <!-- Message Read mode-->
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
                                            <!-- Enables your add-in to activate on multiple selected messages. -->
                                            <SupportsMultiSelect>true</SupportsMultiSelect>
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
                  <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
                </bt:Urls>
                <bt:ShortStrings>
                  <bt:String id="GroupLabel" DefaultValue="Item Multi-select"/>
                  <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
                </bt:ShortStrings>
                <bt:LongStrings>
                  <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane which displays an option to retrieve the subject line of selected messages."/>
                </bt:LongStrings>
            </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

1. Save your changes.

---

> [!NOTE]
> If you turn on the item multi-select feature in your add-in, your add-in will automatically support the [no item context](contextless.md) feature, even if it isn't explicitly configured in the manifest.

## Configure the task pane

Item multi-select relies on the [SelectedItemsChanged](/javascript/api/office/office.eventtype) event to determine when messages are selected or deselected. This event requires a task pane implementation.

1. From the **./src/taskpane** folder, open **taskpane.html**.

1. In the **\<body\>** element, replace the entire **\<main\>** element with the following markup.

    ```html
    <main id="app-body" class="ms-welcome__main">
        <h2 class="ms-font-xl">Retrieve the subject line of multiple messages with one click!</h2>
        <ul id="selected-items"></ul>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. Save your changes.

## Implement a handler for the SelectedItemsChanged event

To alert your add-in when the `SelectedItemsChanged` event occurs, you must register an event handler using the `addHandlerAsync` method.

1. From the **./src/taskpane** folder, open **taskpane.js**.

1. In the `Office.onReady()` callback function, replace the existing code with the following:

    ```javascript
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    
        // Register an event handler to identify when messages are selected.
        Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }
    
          console.log("Event handler added.");
        });
    }
    ```

## Retrieve the subject line of selected messages

Now that you've registered an event handler, you then call the [getSelectedItemsAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getselecteditemsasync-member(1)) method to retrieve the subject line of the selected messages and log them to the task pane. The `getSelectedItemsAsync` method can also be used to get other message properties, such as the item ID, item type (`Message` is the only supported type at this time), and item mode (`Read` or `Compose`).

1. In **taskpane.js**, navigate to the `run` function and insert the following code.

    ```javascript
    // Clear list of previously selected messages, if any.
    const list = document.getElementById("selected-items");
    while (list.firstChild) {
        list.removeChild(list.firstChild);
    }

    // Retrieve the subject line of the selected messages and log it to a list in the task pane.
    Office.context.mailbox.getSelectedItemsAsync(asyncResult => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;      
        }

        asyncResult.value.forEach(item => {
            const listItem = document.createElement("li");
            listItem.textContent = item.subject;
            list.appendChild(listItem);
        });
    });
    ```

1. Save your changes.

## Try it out

1. From a terminal, run the following code in the root directory of your project. This starts the local web server and sideloads your add-in.

    ```command line
    npm start
    ```

    > [!TIP]
    > If your add-in doesn't automatically sideload, follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload it in Outlook.

1. In Outlook, ensure the Reading Pane is enabled. To enable the Reading Pane, see [Use and configure the Reading Pane to preview messages](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0).

1. Navigate to your inbox and choose multiple messages by holding **Ctrl** while selecting messages.

1. Select **Show Taskpane** from the ribbon.

1. In the task pane, select **Run** to view a list of the selected messages' subject lines.

    :::image type="content" source="../images/outlook-multi-select.png" alt-text="A sample list of subject lines retrieved from multiple selected messages.":::

## Item multi-select behavior and limitations

Item multi-select only supports messages within an Exchange mailbox in both read and compose modes. An Outlook add-in only activates on multiple messages if the following conditions are met.

- The messages must be selected from one Exchange mailbox at a time. Non-Exchange mailboxes aren't supported.
- The messages must be selected from one mailbox folder at a time. An add-in doesn't activate on multiple messages if they're located in different folders, unless Conversations view is enabled. For more information, see [Multi-select in conversations](#multi-select-in-conversations).
- An add-in must implement a task pane in order to detect the `SelectedItemsChanged` event.
- The [Reading Pane](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) in Outlook must be enabled. An exception to this is if the item multi-select feature is enabled through the no item context feature in the manifest. To learn more, see [Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md).
- A maximum of 100 messages can be selected at a time.

> [!NOTE]
> Meeting invites and responses are considered messages, not appointments, and can therefore be included in a selection.

### Multi-select in conversations

Item multi-select supports [Conversations view](https://support.microsoft.com/office/0eeec76c-f59b-4834-98e6-05cfdfa9fb07) whether it's enabled on your mailbox or on specific folders. The following table describes expected behaviors when conversations are expanded or collapsed, when the conversation header is selected, and when conversation messages are located in a different folder from the one currently in view.

|Selection|Expanded conversation view|Collapsed conversation view|
|------|------|------|
|**Conversation header is selected**|If the conversation header is the only item selected, an add-in supporting multi-select doesn't activate. However, if other non-header messages are also selected, the add-in will only activate on those and not the selected header.|The behavior differs depending on the Outlook client.<br><br>**Outlook on Windows (classic) and on Mac**:<br>The newest message (that is, the first message in the conversation stack) is included in the message selection.<br><br>If the newest message in the conversation is located in another folder from the one currently in view, the subsequent message in the stack located in the current folder is included in the selection.<br><br>**Outlook on the web and new Outlook on Windows**:<br>All the messages in the conversation stack are selected. This includes messages in the conversation that are located in folders other than the one currently in view.|
|**Multiple selected messages in a conversation stack are located in the same folder as the one currently in view**|All chosen messages in the same conversation are included in the selection.|Not applicable. You must expand the conversation stack to select multiple messages from it.|
|**Multiple selected messages in a conversation stack are located in different folders from the one currently in view** |All chosen messages in the same conversation are included in the selection.|Not applicable. You must expand the conversation stack to select multiple messages from it.|

> [!NOTE]
> On all Outlook clients, you can't select multiple messages that belong to different conversations. If you expand a different conversation while another conversation is expanded, the view of the currently expanded conversation collapses and any selected messages are deselected. However, you can select multiple messages from the same expanded conversation and messages that aren't part of any conversation at the same time.

### Task pane pinning in multi-select add-ins

In Outlook on the web and in new Outlook on Windows, when the task pane of a multi-select add-in is opened, it's automatically pinned to the Outlook client. It remains pinned even when a user switches to a different mail item or selects the **pin** icon from the task pane. The task pane can only be closed by selecting the **Close** button from the task pane.

Conversely, in Outlook on Windows (classic) and on Mac, the task pane isn't automatically pinned and closes when a user switches to a different mail item.

## Next steps

Now that you've enabled your add-in to operate on multiple selected messages, you can extend your add-in's capabilities and further enhance the user experience. Explore performing more complex operations by using the selected messages' item IDs with services, such as [Microsoft Graph](/graph/overview).

## See also

- [Office add-in manifests](../develop/add-in-manifests.md)
- [Call web services from an Outlook add-in](web-services.md)
- [Overview of Microsoft Graph](/graph/overview)
- [Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md)
