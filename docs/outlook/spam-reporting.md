---
title: Implement an integrated spam-reporting add-in
description: Learn how to implement an integrated spam-reporting add-in in Outlook.
ms.date: 08/26/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement an integrated spam-reporting add-in

With the number of unsolicited emails on the rise, security is at the forefront of add-in usage. Currently, partner spam-reporting add-ins are added to the Outlook ribbon, but they usually appear towards the end of the ribbon or in the overflow menu. This makes it harder for users to locate the add-in to report unsolicited emails. In addition to configuring how messages are processed when they're reported, developers also need to complete additional tasks to show processing dialogs or supplemental information to the user.

The integrated spam-reporting feature eases the task of developing individual add-in components from scratch. More importantly, it displays your add-in in a prominent spot on the Outlook ribbon, making it easier for users to locate it and report spam messages. Implement this feature in your add-in to:

- Improve how unsolicited messages are tracked.
- Provide better guidance to users on how to report suspicious messages.
- Enable an organization's security operations center (SOC) or IT administrators to easily perform spam and phishing simulations for educational purposes.

> [!NOTE]
> Integrated spam reporting was introduced in [Mailbox requirement set 1.14](/javascript/api/requirement-sets/outlook/requirement-set-1.14/outlook-requirement-set-1.14). Additional functionality was also added to subsequent requirement sets. To verify that an Outlook client supports these features, see [Supported clients](#supported-clients) and the relevant sections of this article for the features you want to implement.

## Supported clients

The following table identifies which Outlook clients support the integrated spam-reporting feature.

| Client | Status |
| ---- | ---- |
| **Outlook on the web** | Supported\* |
| [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) | Supported\* |
| **classic Outlook on Windows**<br>Version 2404 (Build 17530.15000) | Supported |
| **Outlook on Mac**<br>Version 16.100 (25072537) | Supported |
| **Outlook on Android** | Not available |
| **Outlook on iOS** | Not available |

> [!NOTE]
> \* In Outlook on the web and the new Outlook on Windows, the integrated spam-reporting feature isn't supported for Microsoft 365 consumer accounts.

## Set up your environment

> [!TIP]
> To immediately try out a completed spam-reporting add-in solution, see the [Report spam or phishing emails in Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-spam-reporting) sample.

Complete the [Outlook quick start](../quickstarts/outlook-quickstart-yo.md), which creates an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

To implement the integrated spam-reporting feature in your add-in, you must configure the following in your manifest.

- The runtime used by the add-in. In classic Outlook on Windows, a spam-reporting add-in runs in a [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime). In Outlook on the web and on Mac and in the new Outlook on Windows, a spam-reporting add-in runs in a [browser runtime](../testing/runtimes.md#browser-runtime). For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md).
- The button of the spam-reporting add-in that always appears in a prominent spot on the Outlook ribbon. The following is an example of how the button of a spam-reporting add-in appears on the ribbon of the classic Outlook client on Windows. The ribbon UI may vary depending on the platform the user's Outlook client is running on.

    :::image type="content" source="../images/outlook-spam-ribbon-button.png" alt-text="A sample ribbon button of a spam-reporting add-in.":::

- The preprocessing dialog. When a user selects the add-in button, a dialog appears with **Report** and **Don't Report** options. In this dialog, you can share information about the reporting process or implement input options to get information about a reported message. Input options include checkboxes, radio buttons, and a text box. When a user selects **Report** from the dialog, the [SpamReporting](/javascript/api/office/office.eventtype) event is activated and is then handled by the JavaScript event handler. The following is an example of a preprocessing dialog in classic Outlook on Windows. Note that the appearance of the dialog may vary depending on the platform the user's Outlook client is running on.

    :::image type="content" source="../images/outlook-spam-processing-dialog.png" alt-text="A sample preprocessing dialog of a spam-reporting add-in.":::

Select the tab for the type of manifest you're using.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

> [!NOTE]
> Implementing integrated spam reporting with the unified manifest for Microsoft 365 is currently only available in classic Outlook on Windows. For more information, see the [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).

1. In your preferred code editor, open the add-in project you created.
1. Open the **manifest.json** file.
1. Update the [`"$schema"`](/microsoft-365/extensibility/schema/root#schema-5) property to use the latest version. For more information, see [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).
1. Add the following object to the [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array) array. Note the following about this markup.
   - The [`"minVersion"`](/microsoft-365/extensibility/schema/requirements-extension-element-capabilities#minversion) of the Mailbox requirement set is configured to `"1.14"`. This is the lowest version of the requirement set that supports the integrated spam-reporting feature.
   - The [`"id"`](/microsoft-365/extensibility/schema/extension-runtimes-array#id) of the runtime is set to a unique descriptive name, `"spam_reporting_runtime"`.
   - The [`"code"`](/microsoft-365/extensibility/schema/extension-runtimes-array#code) property has a child `"page"` property that's set to an HTML file and a child `"script"` property that's set to a JavaScript file. You'll create or edit these files in later steps.
   - The [`"lifetime"`](/microsoft-365/extensibility/schema/extension-runtimes-array#lifetime) property is set to `"short"`. This means that the runtime starts when the `SpamReporting` event occurs and shuts down when the event handler completes.
   - The [`"actions"`](/microsoft-365/extensibility/schema/extension-runtimes-array#actions) object specifies the event handler function that runs in the runtime. You'll create this function in a later step.

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.14"
                }
            ]
        },
        "id": "spam_reporting_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/spamreporting.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "onSpamReport",
                "type": "executeFunction"
            }
        ]
    },
    ```

1. Add the following object to the [`"extensions.ribbons"`](/microsoft-365/extensibility/schema/extension-ribbons-array) array. Note the following about this markup.
    - The [`"contexts"`](/microsoft-365/extensibility/schema/extension-ribbons-array#contexts) array contains the `"spamReportingOverride"` string. This prevents the add-in button from appearing at the end of the ribbon or in the overflow section.
    - The [`"tabs"`](/microsoft-365/extensibility/schema/extension-ribbons-array#tabs) array must be specified in an `"extensions.ribbons"` object. However, because the button of a spam-reporting add-in is displayed in a specific spot on the ribbon, only an empty array is specified.
    - The [`"fixedControls"`](/microsoft-365/extensibility/schema/extension-ribbons-array#fixedcontrols) array contains an object that configures the look and functionality of the add-in button on the ribbon. The name of the event handler specified in the [`"actionId"`](/microsoft-365/extensibility/schema/extension-ribbons-array-fixed-control-item#actionid) property must match the value used in the `"id"` property of the object in the `"actions"` array. While the [`"enabled"`](/microsoft-365/extensibility/schema/extension-ribbons-array-fixed-control-item#enabled) property must be specified in the array, its value doesn't affect the functionality of a spam-reporting add-in.
    - The [`"spamPreProcessingDialog"`](/microsoft-365/extensibility/schema/extension-ribbons-array#spampreprocessingdialog) object specifies the information and options that are shown in the preprocessing dialog. While you must specify a [`"title"`](/microsoft-365/extensibility/schema/extension-ribbons-spam-pre-processing-dialog#title) and [`"description"`](/microsoft-365/extensibility/schema/extension-ribbons-spam-pre-processing-dialog#description) for the dialog, you can optionally configure the following properties.
        - The [`"spamNeverShowAgainOption"`](/microsoft-365/extensibility/schema/extension-ribbons-spam-pre-processing-dialog#spamnevershowagainoption) property. It provides the option to suppress the preprocessing dialog for succeeding reports. This functionality is useful in scenarios where you don't need additional information from users. To learn more, see [Suppress the preprocessing dialog](#suppress-the-preprocessing-dialog).
        - The [`"spamReportingOptions"`](/microsoft-365/extensibility/schema/extension-ribbons-spam-pre-processing-dialog-spam-reporting-options?view=m365-app-1.22&preserve-view=true) object. It provides a list of up to five choices that you can configure as radio buttons or checkboxes. This helps a user identify the type of message they're reporting.
        - The [`"spamFreeTextSectionTitle"`](/microsoft-365/extensibility/schema/extension-ribbons-spam-pre-processing-dialog#spamfreetextsectiontitle) property. It provides a text box for the user to add more information about the message they're reporting.
        - The [`"spamMoreInfo"`](/microsoft-365/extensibility/schema/extension-ribbons-spam-pre-processing-dialog-spam-more-info) object. It includes a link in the dialog to provide informational resources to the user.

    ```json
    {
        "contexts": [
            "spamReportingOverride"
        ],
        "tabs": [],
        "fixedControls": [
            {
                "id": "spamReportingButton",
                "type": "button",
                "label": "Report Spam Message",
                "enabled": false,
                "icons": [
                    {
                        "size": 16,
                        "url": "https://localhost:3000/assets/icon-16.png"
                    },
                    {
                        "size": 32,
                        "url": "https://localhost:3000/assets/icon-32.png"
                    },
                    {
                        "size": 80,
                        "url": "https://localhost:3000/assets/icon-80.png"
                    }
                ],
                "supertip": {
                    "title": "Report Spam Message",
                    "description": "Report an unsolicited message."
                },
                "actionId": "onSpamReport"
            }
        ],
        "spamPreProcessingDialog": {
            "title": "Report Spam Message",
            "description": "Thank you for reporting this message.",
            "spamNeverShowAgainOption": false,
            "spamReportingOptions": {
                "title": "Why are you reporting this email?",
                "options": [
                    "Received spam email.",
                    "Received a phishing email.",
                    "I'm not sure this is a legitimate email."
                ],
                "type": "checkbox"
            },
            "spamFreeTextSectionTitle": "Provide additional information, if any:",
            "spamMoreInfo": {
                "text": "Reporting unsolicited messages",
                "url": "https://www.contoso.com/spamreporting"
            }
        }
    },
    ```

1. Save your changes.

# [Add-in only manifest](#tab/xmlmanifest)

Configure the [VersionOverridesV1_1](/javascript/api/manifest/versionoverrides-1-1-mail) node of your add-in only manifest accordingly.

- To run a spam-reporting add-in in Outlook on the web and on Mac and in the new Outlook on Windows, you must specify the HTML file that references or contains the code to handle the spam-reporting event in the `resid` attribute of the [Runtime](/javascript/api/manifest/runtime) element.
- To run a spam-reporting add-in in classic Outlook on Windows, you must specify the JavaScript file that contains the code to handle the spam-reporting event in the [Override](/javascript/api/manifest/override) child element of the `<Runtime>` element.
- To activate the add-in in the Outlook ribbon and prevent it from appearing at the end of the ribbon or in the overflow section, set the `xsi:type` attribute of the `<ExtensionPoint>` element to [ReportPhishingCommandSurface](/javascript/api/manifest/extensionpoint#reportphishingcommandsurface).
- To customize the ribbon button and preprocessing dialog, you must define the [ReportPhishingCustomization](/javascript/api/manifest/reportphishingcustomization) node.
  - To configure the ribbon button, set the `xsi:type` attribute of the [Control](/javascript/api/manifest/control-button) element to `Button`. Then, set the `xsi:type` attribute of the [Action](/javascript/api/manifest/action) child element to `ExecuteFunction` and specify the name of the spam-reporting event handler in its `<FunctionName>` child element.
  - To customize the preprocessing dialog, configure the [PreProcessingDialog](/javascript/api/manifest/preprocessingdialog) element of your manifest. While the dialog must have a title and description, you can optionally include the following elements.
    - A list of choices to help a user identify the type of message they're reporting. You can configure these options as radio buttons or checkboxes. To learn more, see [ReportingOptions element](/javascript/api/manifest/reportingoptions).
    - A text box for the user to provide additional information about the message they're reporting. To learn how to implement a text box, see [FreeTextLabel element](/javascript/api/manifest/preprocessingdialog#child-elements).
    - Custom text and URL to provide informational resources to the user. To learn how to personalize these elements, see [MoreInfo element](/javascript/api/manifest/moreinfo).

      > [!NOTE]
      > Depending on the Outlook client, the custom text specified in the `<MoreInfoText>` element appears before the URL that's provided in the `<MoreInfoUrl>` element or as link text for the URL. For more information, see [MoreInfoText](/javascript/api/manifest/moreinfo#moreinfotext).
    - A "Don't show me this message again" checkbox to prevent the preprocessing dialog from appearing again. To learn how to implement this option, see [Suppress the preprocessing dialog](#suppress-the-preprocessing-dialog).

The following is an example of a `<VersionOverrides>` node configured for spam reporting.

1. In your preferred code editor, open the add-in project you created.
1. Open the **manifest.xml** file located at the root of your project.
1. Select the entire `<VersionOverrides>` node (including the open and close tags) and replace it with the following code.

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.14">
            <bt:Set Name="Mailbox"/>
          </bt:Sets>
        </Requirements>
        <Hosts>
          <Host xsi:type="MailHost">
            <Runtimes>
                <!-- References the HTML file that links to the spam-reporting event handler.
                     This is used by Outlook on the web and on Mac, and new Outlook on Windows. -->
              <Runtime resid="WebViewRuntime.Url">
                <!-- References the JavaScript file that contains the spam-reporting event handler. This is used by classic Outlook on Windows. -->
                <Override type="javascript" resid="JSRuntime.Url"/>
              </Runtime>
            </Runtimes>
            <DesktopFormFactor>
              <FunctionFile resid="WebViewRuntime.Url"/>
              <!-- Implements the integrated spam-reporting feature in the add-in. -->
              <ExtensionPoint xsi:type="ReportPhishingCommandSurface">
                <ReportPhishingCustomization>
                  <!-- Configures the ribbon button. -->
                  <Control xsi:type="Button" id="spamReportingButton">
                    <Label resid="spamButton.Label"/>
                    <Supertip>
                      <Title resid="spamButton.Label"/>
                      <Description resid="spamSuperTip.Text"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>onSpamReport</FunctionName>
                    </Action>
                  </Control>
                  <!-- Configures the preprocessing dialog. -->
                  <PreProcessingDialog>
                    <Title resid="PreProcessingDialog.Label"/>
                    <Description resid="PreProcessingDialog.Text"/>
                    <ReportingOptions>
                      <Title resid="OptionsTitle.Label"/>
                      <Option resid="Option1.Label"/>
                      <Option resid="Option2.Label"/>
                      <Option resid="Option3.Label"/>
                    </ReportingOptions>
                    <FreeTextLabel resid="FreeText.Label"/>
                    <MoreInfo>
                      <MoreInfoText resid="MoreInfo.Label"/>
                      <MoreInfoUrl resid="MoreInfo.Url"/>
                    </MoreInfo>
                  </PreProcessingDialog>
                 <!-- Identifies the runtime to be used. This is also referenced by the Runtime element. -->
                  <SourceLocation resid="WebViewRuntime.Url"/>
                </ReportPhishingCustomization> 
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
            <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html"/>
            <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/spamreporting.js"/>
            <bt:Url id="MoreInfo.Url" DefaultValue="https://www.contoso.com/spamreporting"/>
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="spamButton.Label" DefaultValue="Report Spam Message"/>
            <bt:String id="PreProcessingDialog.Label" DefaultValue="Report Spam Message"/>
            <bt:String id="OptionsTitle.Label" DefaultValue="Why are you reporting this email?"/>
            <bt:String id="FreeText.Label" DefaultValue="Provide additional information, if any:"/>
            <bt:String id="MoreInfo.Label" DefaultValue="Reporting unsolicited messages"/>
            <bt:String id="Option1.Label" DefaultValue="Received spam email."/>
            <bt:String id="Option2.Label" DefaultValue="Received a phishing email."/>
            <bt:String id="Option3.Label" DefaultValue="I'm not sure this is a legitimate email."/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="spamSuperTip.Text" DefaultValue="Report an unsolicited message."/>
            <bt:String id="PreProcessingDialog.Text" DefaultValue="Thank you for reporting this message."/>
          </bt:LongStrings>
        </Resources>
      </VersionOverrides>
    </VersionOverrides>
    ```

1. Save your changes.

---

> [!TIP]
> To learn more about manifests for Outlook add-ins, see [Office Add-ins manifest](../develop/add-in-manifests.md).

## Implement the event handler

When your add-in is used to report a message, it generates a `SpamReporting` event, which is then processed by the event handler in the JavaScript file of your add-in. To map the name of the event handler you specified in your manifest to its JavaScript counterpart, you must call [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) in your code.

1. In your add-in project, navigate to the **./src** directory. Then, create a new folder named **spamreporting**.
1. In the **./src/spamreporting** folder, create a new file named **spamreporting.js**.
1. Open the newly created **spamreporting.js** file and add the following JavaScript code.

    ```javascript
    // Handles the SpamReporting event to process a reported message.
    function onSpamReport(event) {
      // TODO - Send a copy of the reported message.

      // TODO - Get the user's responses.

      // TODO - Signal that the spam-reporting event has completed processing.
    }

    // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
    Office.actions.associate("onSpamReport", onSpamReport);
    ```

1. Save your changes.

### Forward a copy of the message and get the preprocessing dialog responses

Your event handler is responsible for processing the reported message. You can configure it to forward information, such as a copy of the message or the options selected by the user in the preprocessing dialog, to an internal system for further investigation.

To efficiently send a copy of the reported message, call the [getAsFileAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getasfileasync-member(1)) method in your event handler. This gets the Base64-encoded EML format of a message, which you can then forward to your internal system.

If you need to keep track of the user's responses to the options and text box in the preprocessing dialog, extract the `options` and `freeText` values from the `SpamReporting` event object. For more information about these properties, see [Office.SpamReportingEventArgs](/javascript/api/outlook/office.spamreportingeventargs).

The following is an example of a spam-reporting event handler that calls the `getAsFileAsync` method and gets the user's responses from the `SpamReporting` event object.

1. In the **spamreporting.js** file, replace its contents with the following code.

    ```javascript
    // Handles the SpamReporting event to process a reported message.
    function onSpamReport(event) {
      // Get the Base64-encoded EML format of a reported message.
      Office.context.mailbox.item.getAsFileAsync({ asyncContext: event }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
          return;
        }

        // Get the user's responses to the options and text box in the preprocessing dialog.
        const spamReportingEvent = asyncResult.asyncContext;
        const reportedOptions = spamReportingEvent.options;
        const additionalInfo = spamReportingEvent.freeText;

        // Run additional processing operations here.

        // TODO - Signal that the spam-reporting event has completed processing.
      });
    }

    // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
    Office.actions.associate("onSpamReport", onSpamReport);
    ```

1. Save your changes.

> [!NOTE]
> To configure single sign-on (SSO) or cross-origin resource sharing (CORS) in your spam-reporting add-in, you must add your add-in and its JavaScript file to a well-known URI. For guidance on how to configure this resource, see [Use single sign-on (SSO) or cross-origin resource sharing (CORS) in your event-based or spam-reporting Outlook add-in](../develop/use-sso-in-event-based-activation.md).

### Signal when the event has been processed

Once the event handler has completed processing the message, it must call the [event.completed](/javascript/api/outlook/office.mailboxevent#outlook-office-mailboxevent-completed-member(1)) method. In addition to signaling to the add-in that the spam-reporting event has been processed, `event.completed` can also be used to do the following:

- Customize a post-processing dialog to show to the user.
- Perform additional operations on the message, such as deleting it from the inbox.
- [Open a task pane and pass information to it](#open-a-task-pane-after-reporting-a-message).

For a list of properties you can include in a JSON object to pass as a parameter to the `event.completed` method, see [Office.SpamReportingEventCompletedOptions](/javascript/api/outlook/office.spamreportingeventcompletedoptions).

> [!NOTE]
> Code added after the `event.completed` call isn't guaranteed to run.

1. In the **spamreporting.js** file, replace its contents with the following code.

    ```javascript
    // Handles the SpamReporting event to process a reported message.
    function onSpamReport(event) {
      // Get the Base64-encoded EML format of a reported message.
      Office.context.mailbox.item.getAsFileAsync({ asyncContext: event }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
          return;
        }

        // Get the user's responses to the options and text box in the preprocessing dialog.
        const spamReportingEvent = asyncResult.asyncContext;
        const reportedOptions = spamReportingEvent.options;
        const additionalInfo = spamReportingEvent.freeText;

        // Run additional processing operations here.

        /**
         * Signals that the spam-reporting event has completed processing.
         * It then moves the reported message to the Junk Email folder of the mailbox, then
         * shows a post-processing dialog to the user. If an error occurs while the message
         * is being processed, the `onErrorDeleteItem` property determines whether the message
         * will be deleted.
         */
        const event = asyncResult.asyncContext;
        event.completed({
          onErrorDeleteItem: true,
          moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
          showPostProcessingDialog: {
            title: "Contoso Spam Reporting",
            description: "Thank you for reporting this message.",
          },
        });
      });
    }

    // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart
    Office.actions.associate("onSpamReport", onSpamReport);
    ```

    > [!NOTE]
    > If you're on classic Outlook on Windows Version 2308 (Build 16724.10000) or later, Outlook on Mac, Outlook on the web, or the new Outlook on Windows, you must use the `moveItemTo` property in the `event.completed` call to specify the folder to which a reported message is moved once it's processed by your add-in. On earlier Outlook builds on Windows that support the integrated spam-reporting feature, you must use the `postProcessingAction` property.

1. Save your changes.

The following is a sample post-processing dialog shown to the user once the add-in completes processing a reported message in classic Outlook on Windows. Note that the appearance of the dialog may vary depending on the platform the user's Outlook client is running on.

:::image type="content" source="../images/outlook-spam-post-processing-dialog.png" alt-text="A sample of a post-processing dialog shown once a reported spam message has been processed by the add-in.":::

> [!TIP]
> As you develop a spam-reporting add-in that will run in classic Outlook on Windows, keep the following in mind.
>
> - Imports aren't currently supported in the JavaScript file that contains the code to handle the spam-reporting event.
> - When the JavaScript function specified in the manifest to handle the `SpamReporting` event runs, code in `Office.onReady()` and `Office.initialize` isn't run. We recommend adding any startup logic needed by the event handler, such as checking the user's Outlook version, to the event handler instead.

## Update the commands HTML file

1. In the **./src/commands** folder, open **commands.html**.
1. Immediately before the closing **head** tag (`</head>`), add the following **script** entry.

    ```html
    <script type="text/javascript" src="../spamreporting/spamreporting.js"></script>    
    ```

1. Save your changes.

## Update the webpack config settings

1. From the root directory of your add-in project, open the **webpack.config.js** file.

1. Locate the `plugins` array within the `config` object and add this new object to the beginning of the array.

    ```js
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "./src/spamreporting/spamreporting.js",
          to: "spamreporting.js",
        },
      ],
    }),
    ```

1. Save your changes.

## Test and validate your add-in

1. [Sideload](sideload-outlook-add-ins-for-testing.md) the add-in in a supported Outlook client.
1. Choose a message from your inbox, then select the add-in's button from the ribbon.
1. In the preprocessing dialog, choose a reason for reporting the message and add information about the message, if configured. Then, select **Report**.
1. (Optional) In the post-processing dialog, select **OK**.

## Suppress the preprocessing dialog

> [!NOTE]
> The "Don't show me this message again" option was introduced in [requirement set 1.15](/javascript/api/requirement-sets/outlook/requirement-set-1.15/outlook-requirement-set-1.15). Learn more about its [supported clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support).

Depending on your scenario, you might not need a user to provide additional information about a message they're reporting. If the preprocessing dialog of your spam-reporting add-in only provides information to the user, you can choose to include a "Don't show me this message again" option in the dialog.

:::image type="content" source="../images/spam-reporting-suppress-dialog.png" alt-text="The 'Don't show me this message again' option in the preprocessing dialog.":::

To implement this in your add-in, you must configure its manifest. The element or property to be configured differs depending on the type of manifest your add-in uses.

- **Unified manifest for Microsoft 365**: Set the `"extensions.ribbons.spamPreProcessingDialog.spamNeverShowAgainOption"` property to `true`. The following code is an example.

    ```json
    "extensions": [
        ...
        "ribbons": [
            {
                ...
                "spamPreProcessingDialog": {
                    "title": "Report Spam Message",
                    "description": "Thank you for reporting this message.",
                    "spamNeverShowAgainOption": true,
                    "spamMoreInfo": {
                        "text": "Reporting unsolicited messages",
                        "url": "https://www.contoso.com/spamreporting"
                    }
                }
            }
        ],
        ...
    ]
    ```

- **Add-in only manifest**: In the [PreProcessingDialog](/javascript/api/manifest/preprocessingdialog) element, set the [NeverShowAgainOption](/javascript/api/manifest/preprocessingdialog#child-elements) child element to `true`. The following code is an example.

    ```xml
    ...
        <PreProcessingDialog>
          <Title resid="PreProcessingDialog.Label"/>
          <Description resid="PreProcessingDialog.Text"/>
          <NeverShowAgainOption>true</NeverShowAgainOption>
          <MoreInfo>
            <MoreInfoText resid="MoreInfo.Label"/>
            <MoreInfoUrl resid="MoreInfo.Url"/>
          </MoreInfo>
        </PreProcessingDialog>
    ...
    ```

Note the following behaviors when implementing this option in your add-in.

- The "Don't show me this message again" option appears in the preprocessing dialog only if there are no user input options configured, such as reporting options or a text box. If your manifest also specifies user input options, they'll be shown in the dialog instead.
- The option to suppress the preprocessing dialog is applied on a per-machine and per-platform basis.
- After suppressing the preprocessing dialog, if a user reports a message and has the [Reading Pane](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) turned on, a progress notification is shown on the message while it's being processed. If the Reading Pane is turned off, no progress notification is shown.

    :::image type="content" source="../images/spam-reporting-progress-notification.png" alt-text="The progress notification shown after reporting a message while the Reading Pane is turned on.":::

    > [!NOTE]
    > If you implement the "Don't show me this message again" option without configuring a post-processing dialog, a dialog with the following generic message is shown to the user after processing: "Thank you for reporting this message." It's good practice to configure a post-processing dialog to appear after a reported message is processed.

### Reenable the preprocessing dialog

To reenable the preprocessing dialog in classic Outlook on Windows after selecting the "Don't show me this message again" option, perform the following steps.

1. Open the **Registry Editor** as an administrator.
1. Navigate to **HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\WebExt\SpamDialog**.
1. Locate your spam-reporting add-in using its ID specified in the manifest.
1. Delete the entry of your add-in from the registry.

The preprocessing dialog will appear the next time a message is reported.

## Open a task pane after reporting a message

> [!NOTE]
> The option to implement a task pane from the `event.completed` method was introduced in [requirement set 1.15](/javascript/api/requirement-sets/outlook/requirement-set-1.15/outlook-requirement-set-1.15). Learn more about its [supported clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support).

Instead of a post-processing dialog, you can implement a task pane to open after a user reports a message. For example, you can use the task pane to show additional information based on the user's input in the preprocessing dialog. Similar to the post-processing dialog, the task pane is implemented through the add-in's `event.completed` call.

> [!NOTE]
> If both the post-processing dialog and task pane capabilities are configured in the `event.completed` call, the task pane is shown instead of the dialog.

To configure a task pane to open after a message is reported, you must specify the ID of the task pane in the [commandId](/javascript/api/outlook/office.spamreportingeventcompletedoptions#outlook-office-spamreportingeventcompletedoptions-commandid-member) option of the `event.completed` call. The ID must match the task pane's ID in the manifest. The location of the ID varies depending on the manifest your add-in uses.

- **Unified manifest for Microsoft 365**: The `"id"` property of the [`"extensions.ribbons.tabs.groups.controls"`](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item#id) object that represents the task pane.
- **Add-in only manifest**: The `id` attribute of the [Control](/javascript/api/manifest/control) element that represents the task pane.

If you need to pass information to the task pane, specify any JSON data in the [contextData](/javascript/api/outlook/office.spamreportingeventcompletedoptions#outlook-office-spamreportingeventcompletedoptions-contextdata-member) option of the `event.completed` call. To retrieve the value of the `contextData` option, you must call [Office.context.mailbox.item.getInitializationContextAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getinitializationcontextasync-member(1)) in the JavaScript implementation of your task pane.

> [!IMPORTANT]
> To ensure that the task pane of the spam-reporting add-in opens and receives context data after a message is reported, you must set the `moveItemTo` option of the `event.completed` call to `Office.MailboxEnums.MoveSpamItemTo.NoMove`.

The following code is an example.

```javascript
...
    event.completed({
      commandId: "msgReadOpenPaneButton",
      contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
      moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove
    });
```

If another task pane is open or pinned at the time a message is reported, the open task pane is closed then the task pane of the spam-reporting add-in is shown.

## Review feature behavior and limitations

As you develop and test the integrated spam-reporting feature in your add-in, be mindful of its characteristics and limitations.

- In Outlook on the web and on Windows (new and classic), an integrated spam-reporting add-in replaces the native **Report** button in the Outlook ribbon. If multiple spam-reporting add-ins are installed, they will all appear in the **Report** section of the ribbon.

    :::image type="content" source="../images/outlook-spam-replace-button.png" alt-text="A sample integrated spam-reporting add-in that replaces the Report button in the Outlook ribbon.":::

- A spam-reporting add-in can run for a maximum of five minutes once it's activated. Any processing that occurs beyond five minutes will cause the add-in to time out. If the add-in times out, a dialog will be shown to the user to notify them of this.

  :::image type="content" source="../images/outlook-spam-timeout-dialog.png" alt-text="The dialog shown when a spam-reporting add-in times out.":::

- In classic Outlook on Windows, a spam-reporting add-in can be used to report a message even if the Reading Pane of the Outlook client is turned off. In Outlook on the web, on Mac, and in new Outlook on Windows, the spam-reporting add-in can be used if the Reading Pane is turned on or the message to be reported is open in another window.
- Only one message can be reported at a time. If you select multiple messages to report, the button of the spam-reporting add-in becomes unavailable.
- In classic Outlook on Windows, only one reported message can be processed at a time. If a user attempts to report another message while the previous one is still being processed, a dialog will be shown to notify them of this.

  :::image type="content" source="../images/outlook-spam-report-error.png" alt-text="The dialog shown when the user attempts to report another message while the previous one is still being processed.":::

  This doesn't apply to Outlook on the web or on Mac, or to the new Outlook on Windows. In these Outlook clients, a user can report a message from the Reading Pane and can simultaneously report each message that's open in a separate window.

- The add-in can still process the reported message even if the user navigates away from the selected message. In Outlook on Mac, this is only supported if a user reports a message while it's open in a separate window. If the user reports a message while viewing it from the Reading Pane and then navigates away from it, the reporting process is terminated.
- In Outlook on the web and the new Outlook on Windows, if a message is reported while it's open in a separate window, a post-processing dialog isn't shown to the user.
- The buttons that appear in the preprocessing and post-processing dialogs aren't customizable. Additionally, the text and buttons in the timeout and ongoing report dialogs can't be modified.
- The integrated spam-reporting and [event-based activation](../develop/event-based-activation.md) features must use the same runtime. Multiple runtimes aren't currently supported in Outlook. To learn more about runtimes, see [Runtimes in Office Add-ins](../testing/runtimes.md).
- A spam-reporting add-in only implements [function commands](../design/add-in-commands.md#types-of-add-in-commands). A task pane command can't be assigned to the spam-reporting button on the ribbon. If you want to implement a task pane separate from the reporting functionality of your add-in, you must configure it in your manifest as follows:
  - **Add-in only manifest**: Include the [Action element](/javascript/api/manifest/action#xsitype-is-showtaskpane) in the manifest and set its `xsi:type` attribute to `ShowTaskpane`.
  - **Unified manifest for Microsoft 365**: Configure a task pane object in the `"extensions.runtimes"` and `"extensions.ribbons"` arrays. For guidance, see the "Add a task pane command" section of [Create add-in commands with the unified manifest for Microsoft 365](../develop/create-addin-commands-unified-manifest.md#add-a-task-pane-command).

  Note that a separate button to activate the task pane will be added to the ribbon, but it won't appear in the dedicated spam-reporting area of the ribbon. If you want to configure a task pane to open after a message is reported, see [Open a task pane after reporting a message](#open-a-task-pane-after-reporting-a-message).

## Troubleshoot your add-in

As you develop your spam-reporting add-in, you might need to troubleshoot issues, such as your add-in not loading. For guidance on how to troubleshoot a spam-reporting add-in, see [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md).

## See also

- [Office Add-ins manifest](../develop/add-in-manifests.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
- [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md)
- [ReportPhishingCommandSurface extension point](/javascript/api/manifest/extensionpoint#reportphishingcommandsurface)
- [Office.MessageRead.getAsFileAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getasfileasync-member(1))
- [Office.MailboxEnums.MoveSpamItemTo](/javascript/api/outlook/office.mailboxenums.movespamitemto)
- [Office.SpamReportingEventArgs](/javascript/api/outlook/office.spamreportingeventargs)
- [Office.SpamReportingEventCompletedOptions](/javascript/api/outlook/office.spamreportingeventcompletedoptions)
