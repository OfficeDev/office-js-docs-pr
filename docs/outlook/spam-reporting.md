---
title: Implement an integrated spam reporting add-in (preview)
description: Learn how to implement an integrated spam reporting add-in in Outlook.
ms.date: 07/14/2023
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement an integrated spam reporting add-in (preview)

With the number of unsolicited emails on the rise, security is at the forefront of add-in usage. Currently, partner spam reporting add-ins are added to the Outlook ribbon, but they usually appear towards the end of the ribbon or in the overflow section. This makes it harder for users to locate the add-in to report unsolicited emails. In addition to configuring how messages are processed when they're reported, developers also need to complete additional tasks to show processing dialogs or supplemental information to the user.

The integrated spam reporting feature eases the task of developing individual add-in components from scratch. More importantly, it displays your add-in in a prominent spot on the Outlook ribbon, making it easier for users to locate and report spam messages. Implement this feature in your add-in to:

- Improve how unsolicited messages are tracked.
- Provide better guidance to users on how to report suspicious messages.
- Enable an organization's security operations center (SOC) or IT administrators to easily perform spam and phishing simulations for educational purposes.

> [!IMPORTANT]
> Features in preview shouldn't be used in production add-ins. We invite you to try out this feature in test or development environments and welcome feedback on your experience through GitHub (see the **Feedback** section at the end of this page).

## Preview the spam reporting feature

To preview the spam reporting feature in Outlook on Windows, you must install Version 2307 (Build 16626.10000) or later. Then, join the [Microsoft 365 Insider program](https://insider.microsoft365.com/join/Windows) and select the **Beta Channel** option to access Office beta builds.

Outlook on Windows includes a local copy of the production and beta versions of Office.js instead of loading from the content delivery network (CDN). By default, the local production copy of the API is referenced. To reference the local beta copy of the API, you must configure your computer's registry. Once you've set up your Outlook client, configure the registry as follows:

1. In the registry, navigate to `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`. If the key doesn't exist, create it.
1. Create an entry named `EnableBetaAPIsInJavaScript` and set its value to `1`.

    ![The EnableBetaAPIsInJavaScript registry value is set to 1."](../images/outlook-beta-registry-key.png)

## Set up your environment

Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator), which creates an add-in project with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Configure the manifest

To implement the integrated spam reporting feature in your add-in, you must configure the [VersionOverridesV1_1](/javascript/api/manifest/versionoverrides-1-1-mail) node of your manifest accordingly.

- In Outlook on Windows, an add-in that implements the spam reporting feature runs in a [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime). As such, you must specify the JavaScript file that contains the code to handle the spam reporting event in the [Override](/javascript/api/manifest/override) child element of the [Runtime](/javascript/api/manifest/runtime) element.
- To activate the add-in in the Outlook ribbon and prevent it from appearing at the end of the ribbon or in the overflow section, set the `xsi:type` attribute of the **\<ExtensionPoint\>** element to [ReportPhishingCommandSurface](/javascript/api/manifest/extensionpoint?view=outlook-js-preview&preserve-view=true#reportphishingcommandsurface-preview).
- To customize the ribbon button and pre-processing dialog, you must define the [ReportPhishingCustomization](/javascript/api/manifest/reportphishingcustomization?view=outlook-js-preview&preserve-view=true) node.
  - A user reports an unsolicited message through the add-in's button in the ribbon. The button shows the pre-processing dialog to the user and activates the [SpamReporting](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) event, which is then handled by the JavaScript event handler. To configure the ribbon button, set the `xsi:type` attribute of the [Control](/javascript/api/manifest/control-button) element to `Button`. Then, set the `xsi:type` attribute of the [Action](/javascript/api/manifest/action) child element to `ExecuteFunction` and specify the name of the spam reporting event handler in its **\<FunctionName\>** child element. A spam reporting add-in can only implement [function commands](../design/add-in-commands.md#types-of-add-in-commands).

    :::image type="content" source="../images/outlook-spam-ribbon-button.png" alt-text="A sample ribbon button of a spam reporting add-in.":::

  - The pre-processing dialog is shown to a user when they report a message. In the dialog, you can share additional guidance on the reporting process and include additional options for the user to provide more information about the message being reported. To customize the title, description, and options of the pre-processing dialog, you must include the [PreProcessingDialog](/javascript/api/manifest/preprocessingdialog?view=outlook-js-preview&preserve-view=true) element in your manifest.

    :::image type="content" source="../images/outlook-spam-processing-dialog.png" alt-text="A sample pre-processing dialog of a spam reporting Outlook add-in.":::

The following is an example of a **\<VersionOverrides\>** node configured for spam reporting.

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
          <Runtime resid="WebViewRuntime.Url">
            <!-- References the JavaScript file that contains the spam reporting event handler. This is used by Outlook on Windows. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="WebViewRuntime.Url"/>
          <!-- Implements the integrated spam reporting feature in the add-in. -->
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
              <!-- Configures the pre-processing dialog. -->
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
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/commands.js"/>
        <bt:Url id="MoreInfo.Url" DefaultValue="https://www.contoso.com/spamreporting"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="spamButton.Label" DefaultValue="Report Spam Message"/>
        <bt:String id="PreProcessingDialog.Label" DefaultValue="Report Spam Message"/>
        <bt:String id="OptionsTitle.Label" DefaultValue="Why are you reporting this email?"/>
        <bt:String id="FreeText.Label" DefaultValue="Provide additional information, if any:"/>
        <bt:String id="MoreInfo.Label" DefaultValue="To learn more about reporting unsolicited messages, see "/>
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

## Implement the event handler

When your add-in is used to report a message, it generates a [SpamReporting](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) event, which is then processed by the event handler in the JavaScript file of your add-in. To map the name of the event handler you specified in the **\<FunctionName\>** element of your manifest to its JavaScript counterpart, you must call [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) in your code.

Your event handler is responsible for processing the reported message, such as forwarding a copy of the message to an internal system for further investigation. To efficiently send a copy of the reported message, call the [getAsFileAsync](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-getasfileasync-member(1)) method in your event handler. This gets the Base64 encoding of a message, which you can then forward to your internal system.

Once the event handler has completed processing the message, it must call the [event.completed](/javascript/api/office/office.addincommands.event#office-office-addincommands-event-completed-member(1)) method. In addition to signaling to the add-in that the spam reporting event has been processed, `event.completed` can also be used to customize a post-processing dialog to show to the user or perform additional operations on the message, such as deleting it from the inbox. For a list of properties you can include in a JSON object to pass as a parameter to the `event.completed` method, see [Office.AddinCommands.EventCompletedOptions](/javascript/api/office/office.addincommands.eventcompletedoptions).

> [!NOTE]
> Code added after the `event.completed` call isn't guaranteed to run.

The following is an example of a spam reporting event handler.

```javascript
// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  // Gets the Base64 encoding of a reported message.
  Office.context.mailbox.item.getAsFileAsync({ asyncContext: event }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
      return;
    }

    // Run additional processing operations here.

    /**
     * Signals that the spam reporting event has completed processing.
     * It then moves the reported message to the Junk Email folder of the mailbox, then
     * shows a post-processing dialog to the user. If an error occurs while the message
     * is being processed, the `onErrorDeleteItem` property determines whether the message
     * will be deleted.
     */
    const event = asyncResult.asyncContext;
    event.completed({
      onErrorDeleteItem: true,
      postProcessingAction: moveToSpamFolder,
      showPostProcessingDialog: {
        title: "Contoso Spam Reporting",
        description: "Thank you for reporting this message.",
      },
    });
  });
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onSpamReport", onSpamReport);
}
```

The following is a sample post-processing dialog shown to the user once the add-in completes processing a reported message.

:::image type="content" source="../images/outlook-spam-post-processing-dialog.png" alt-text="A sample of a post-processing dialog shown once a reported spam message has been processed by the add-in.":::

> [!TIP]
> As you develop a spam reporting add-in that will run in Outlook on Windows, keep the following in mind.
>
> - Imports aren't currently supported in the JavaScript file that contains the code to handle the spam reporting event.
> - Code included in the `Office.onReady()` and `Office.initialize` functions won't run. You must add any add-in startup logic, such as checking the user's Outlook version, to your event handlers instead.

## Test and validate your add-in

1. [Sideload](sideload-outlook-add-ins-for-testing.md) the add-in in Outlook on Windows.
1. Select a message from your inbox, then select the add-in's button from the ribbon.
1. In the pre-processing dialog, select a reason for reporting the message and add information about the message, if configured. Then, select **Report**.
1. (Optional) In the post-processing dialog, select **OK**.

## Spam reporting behavior and limitations

As you develop and test the spam reporting feature in your add-in, be mindful of its characteristics and limitations.

- A spam reporting add-in can run for a maximum of five minutes once it's activated. Any processing that occurs beyond five minutes will cause the add-in to time out. If the add-in times out, a dialog will be shown to the user to notify them of this.

  :::image type="content" source="../images/outlook-spam-timeout-dialog.png" alt-text="The dialog shown when a spam reporting add-in times out.":::

- A spam reporting add-in can be used to report a message even if the Reading Pane of the Outlook client is turned off.
- Only one message can be reported at a time. If a user attempts to report another message while the previous one is still being processed, a dialog will be shown to them to notify them of this.

  :::image type="content" source="../images/outlook-spam-report-error.png" alt-text="The dialog shown when the user attempts to report another message while the previous one is still being processed.":::

- The add-in can still process the reported message even if the user navigates away from the selected message.
- The buttons that appear in the pre- and post-processing dialogs aren't customizable. Additionally, the text and buttons in the timeout and ongoing report dialogs can't be modified.

## See also

- [Office Add-ins manifest](../develop/add-in-manifests.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
