---
title: Activate your Outlook add-in without the Reading Pane enabled or a message selected (preview)
description: Learn how to activate your Outlook add-in without enabling the Reading Pane or first selecting a message.
ms.date: 04/21/2023
ms.topic: how-to
ms.localizationpriority: medium
---

# Activate your Outlook add-in without the Reading Pane enabled or a message selected (preview)

With a simple manifest configuration, Outlook add-ins created for the Message Read surface can now activate a task pane without the Reading Pane enabled or a message first selected from the mailbox. Follow the walkthrough to learn more and unlock additional capabilities for your add-in. For example, you can enable your users to access content from different data sources, such as OneDrive or a customer relationship management (CRM) system, directly from their Outlook client.

> [!IMPORTANT]
> Features in preview shouldn't be used in production add-ins. We invite you to test this feature in test or development environments and welcome feedback on your experience through GitHub (see the **Feedback** section at the end of this page).

> [!NOTE]
> This feature isn't currently supported in the [Unified Microsoft 365 manifest (preview)](../develop/json-manifest-overview.md), but the team is working on making this available.

## Prerequisites to preview the feature

To preview this feature, install Outlook on Windows, starting with Version 2304 (Build 16313.10000). Once installed, join the [Office Insider program](https://insider.office.com/join/windows) and select the **Beta Channel** option to access Office beta builds.

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) to create an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

To enable your add-in to activate with the Reading Pane turned off or without a message selected, you must add the [SupportsNoItemContext](/javascript/api/manifest/action?view=outlook-js-preview&preserve-view=true#supportsnoitemcontext-preview) child element to the **\<Action\>** element and set its value to `true`. As this feature can only be implemented with a task pane in Message Read mode, the following elements must also be configured.

- The `xsi:type` attribute value of the **\<ExtensionPoint\>** element must be set to `MessageReadCommandSurface`.
- The `xsi:type` attribute value of the **\<Action\>** element must be set to `ShowTaskpane`.

1. In your preferred code editor, open the Outlook quick start project you created.

1. Open the **manifest.xml** file located at the root of the project.

1. Select the entire **\<VersionOverrides\>** node and replace it with the following XML.

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
                                            <!-- Enables your add-in to activate without the Reading Pane enabled or a message selected. -->
                                            <SupportsNoItemContext>true</SupportsNoItemContext>
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
                  <bt:String id="GroupLabel" DefaultValue="Test walkthrough"/>
                  <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
                </bt:ShortStrings>
                <bt:LongStrings>
                  <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a task pane."/>
                </bt:LongStrings>
            </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

1. Save your changes.

## Configure the task pane

1. In your project, navigate to the **taskpane** folder, then open **taskpane.html**.
1. Replace the entire **\<body\>** element with the following markup.

    ```html
    <body class="ms-font-m ms-welcome ms-Fabric">
        <header class="ms-welcome__header ms-bgColor-neutralLighter">
            <img width="90" height="90" src="../../assets/logo-filled.png" alt="logo" title="Add-in logo" />
            <h1 class="ms-font-su">Activate your add-in without enabling the Reading Pane or selecting a message</h1>
        </header>
        <section id="sideload-msg" class="ms-welcome__main">
            <h2 class="ms-font-xl">Please <a target="_blank" href="https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing">sideload</a> your add-in to see app body.</h2>
        </section>
        <main id="app-body" class="ms-welcome__main" style="display: none;">
            <ul class="ms-List ms-welcome__features">
                <li class="ms-ListItem">
                    <i class="ms-Icon ms-Icon--CheckList ms-font-xl"></i>
                    <span class="ms-font-m">Item multi-select is automatically enabled when the <b>SupportsNoItemContext</b> manifest element is set to <code>true</code>. You can test this by selecting multiple messages in Outlook, then choosing <b>Show Taskpane</b> from the ribbon.</span>
                </li>
                <li class="ms-ListItem">
                    <i class="ms-Icon ms-Icon--Pin ms-font-xl"></i>
                    <span class="ms-font-m">Support to pin the task pane is also automatically enabled. You can test this by selecting the <b>pin</b> icon from the top right corner of the task pane.</span>
                </li>
                <li class="ms-ListItem">
                    <i class="ms-Icon ms-Icon--DockRight ms-font-xl"></i>
                    <span class="ms-font-m">This feature can only be implemented with a task pane.</span>
                </li>
                <li class="ms-ListItem">
                    <i class="ms-Icon ms-Icon--Design ms-font-xl"></i>
                    <span class="ms-font-m">Implement your scenario using this feature today! For example, enable your users to access content from different data sources, such as OneDrive or your customer relationship management (CRM) system, without first selecting a message.</span>
                </li>
            </ul>
        </main>
    </body>
    ```

1. Save your changes.

## Update the task pane JavaScript file

1. From the **taskpane** folder, open **taskpane.js**.
1. Navigate to the `Office.onReady` function and replace its contents with the following code.

    ```javascript
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
    }
    ```

1. Save your changes.

## Try it out

1. From a terminal, run the following code in the root directory of your project. This starts the local web server and sideloads your add-in.

    ```command line
    npm start
    ```

    > [!TIP]
    > If your add-in doesn't automatically sideload, follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md?tabs=windows#sideload-manually) to manually sideload it in Outlook.

1. Navigate to your inbox and do one of the following:

    - Turn off your Reading Pane. For guidance, see the "Turn on, turn off, or move the Reading Pane" section of [Use and configure the Reading Pane to preview messages](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0).
    - Deselect a message, if applicable. To deselect a message, hold the **Ctrl** key and select the message.

1. Select **Show Taskpane** from the ribbon.

1. Explore and test the suggestions listed in the task pane.

## Support for the item multi-select and pinnable task pane features

When the **\<SupportsNoItemContext\>** element in the manifest is set to `true`, it automatically enables the [item multi-select](item-multi-select.md) and [pinnable task pane](pinnable-taskpane.md) features, even if these features aren't explicitly configured in the manifest.

## See also

- [Activate your Outlook add-in on multiple messages (preview)](item-multi-select.md)
- [Implement a pinnable task pane in Outlook](pinnable-taskpane.md)
