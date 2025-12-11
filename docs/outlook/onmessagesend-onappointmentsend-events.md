---
title: Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts
description: Learn about the Smart Alerts implementation and how it handles the OnMessageSend and OnAppointmentSend events in your event-based Outlook add-in.
ms.date: 12/04/2025
ms.topic: concept-article
ms.localizationpriority: medium
---

# Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts

The `OnMessageSend` and `OnAppointmentSend` events take advantage of Smart Alerts, which allows you to run logic after a user selects **Send** in their Outlook message or appointment. With Smart Alerts, users of your add-in can take the opportunity to improve the content of their email, add a missing sensitivity label, or include an important recipient in a meeting invite.

Smart Alerts is available through the event-based activation feature. To understand how to configure your add-in to use this feature, use other available events, debug your add-in, and more, see [Activate add-ins with events](../develop/event-based-activation.md).

> [!NOTE]
> The `OnMessageSend` and `OnAppointmentSend` events were introduced in [requirement set 1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12). Additional functionality and customization options were also added to subsequent requirement sets. To verify that your Outlook client supports these events and features, see [Supported clients and platforms](#supported-clients-and-platforms) and the specific sections in the [walkthrough](smart-alerts-onmessagesend-walkthrough.md) that describe the features you want to implement.

## Supported clients and platforms

The following table lists supported client-server combinations for the Smart Alerts feature, including the minimum required Exchange Server Cumulative Update where applicable. Excluded combinations aren't supported.

|Client|Exchange Online|Exchange Server Subscription Edition (SE)|Exchange 2019 on-premises (Cumulative Update 12 or later)|Exchange 2016 on-premises (Cumulative Update 22 or later) |
|-----|-----|-----|-----|-----|
|**Web browser (modern UI)**|Yes|Not applicable|Not applicable|Not applicable|
|[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)|Yes|Not applicable|Not applicable|Not applicable|
|**Windows** (classic)<br>Version 2206 (Build 15330.20196) or later|Yes|Yes|Yes|Yes|
|**Mac**<br>Version 16.65 (22082700) or later|Yes|Not applicable|Not applicable|Not applicable|
|**Android**|Not applicable|Not applicable|Not applicable|Not applicable|
|**iOS**|Not applicable|Not applicable|Not applicable|Not applicable|

> [!IMPORTANT]
> Enhancements to the Smart Alerts feature were introduced in later requirement sets. Because of this, the minimum supported Outlook version and build may vary. The requirement sets and versions needed for a feature are mentioned in the [walkthrough](smart-alerts-onmessagesend-walkthrough.md) and applicable sections of this article.

## Try out Smart Alerts in an event-based add-in

To see Smart Alerts in action, try out the [walkthrough](smart-alerts-onmessagesend-walkthrough.md). You'll create an add-in that checks whether a document or picture is attached to a message before it's sent. In addition to implementing a basic Smart Alerts add-in, you'll also learn about the following functionalities to further enhance the user experience of your add-in.

- [Customize the Smart Alerts dialog message using Markdown](smart-alerts-onmessagesend-walkthrough.md#implement-event-handling)
- [Customize the text of a dialog button](smart-alerts-onmessagesend-walkthrough.md#customize-the-text-and-functionality-of-a-button-in-the-dialog-optional)
- [Open a task pane or run a function from the dialog](smart-alerts-onmessagesend-walkthrough.md#customize-the-text-and-functionality-of-a-button-in-the-dialog-optional)
- [Override the send mode option at runtime](smart-alerts-onmessagesend-walkthrough.md#override-the-send-mode-option-at-runtime-optional)
- [Programmatically send the mail item once it meets your add-in's conditions](smart-alerts-onmessagesend-walkthrough.md#programmatically-send-the-item-from-the-task-pane-optional)

## Smart Alerts feature behavior and scenarios

The following sections include guidance on the send mode options and the behavior of the feature in certain scenarios.

### Available send mode options

When you configure your add-in to respond to the `OnMessageSend` or `OnAppointmentSend` event, you must include the send mode property in the manifest. Its markup varies depending on the type of manifest your add-in uses.

- **Add-in only manifest**: Set the **SendMode** property of the [LaunchEvent](/javascript/api/manifest/launchevent) element.
- **Unified manifest for Microsoft 365**: Set the `"sendMode"` option of the event object in the [`"autoRunEvents"`](/microsoft-365/extensibility/schema/element-extensions#autorunevents) array.

If the conditions implemented by your add-in aren't met or your add-in is unavailable when the event occurs, a dialog is shown to the user to alert them that additional actions may be needed before the mail item can be sent. The send mode property determines the options available to the user in the dialog.

The following table lists the available send mode options.

|Send mode option canonical name|Add-in only manifest name|Unified manifest for Microsoft 365 name|
|-----|-----|-----|
|**prompt user**|`PromptUser`|promptUser|
|**soft block**|`SoftBlock`|softBlock|
|**block**|`Block`|block|

> [!TIP]
> Starting in [Mailbox requirement set 1.14](/javascript/api/requirement-sets/outlook/requirement-set-1.14/outlook-requirement-set-1.14), your add-in can now override its send mode option at runtime. To learn more, see [Override the send mode option at runtime (optional)](smart-alerts-onmessagesend-walkthrough.md#override-the-send-mode-option-at-runtime-optional).

#### prompt user

If the item doesn't meet the add-in's conditions, the user can choose **Send Anyway** in the alert, or address the issue then try to send the item again. If the add-in is taking a long time to process the item, the user will be prompted with the option to stop running the add-in and choose **Send Anyway**. In the event the add-in is unavailable (for example, there's an error loading the add-in), the item will be sent.

:::image type="content" source="../images/outlook-smart-alerts-prompt-user.png" alt-text="The prompt user dialog with the Send Anyway and Don't Send options.":::

Use the **prompt user** option in your add-in if one of the following applies.

- The condition checked by the add-in isn't mandatory, but is nice to have in the message or appointment being sent.
- You'd like to recommend an action and allow the user to decide whether they want to apply it to the message or appointment being sent.

Some scenarios where the **prompt user** option is applied include suggesting to tag the message or appointment as low or high importance and recommending to apply a color category to the item.

#### soft block

Default option if the send mode property of your manifest isn't configured. The user is alerted that the item they're sending doesn't meet the add-in's conditions and they must address the issue before trying to send the item again. However, if the add-in is unavailable (for example, there's an error loading the add-in), the item will be sent.

:::image type="content" source="../images/outlook-smart-alerts-soft-hard-block.png" alt-text="The soft block dialog with the Don't Send option.":::

Use the **soft block** option in your add-in when you want a condition to be met before a message or appointment can be sent, but you don't want the user to be blocked from sending the item if the add-in is unavailable. Sample scenarios where the **soft block** option is used include prompting the user to set a message or appointment's importance level and checking that the appropriate signature is applied before the item is sent.

#### block

The item isn't sent if any of the following situations occur.

- The item doesn't meet the add-in's conditions.
- The add-in is unable to connect to the server.
- There's an error loading the add-in.

:::image type="content" source="../images/outlook-smart-alerts-soft-hard-block.png" alt-text="The block dialog with the Don't Send option.":::

Use the **block** option if the add-in's conditions are mandatory, even if the add-in is unavailable. For example, the **block** option is ideal when users are required to apply a sensitivity label to a message or appointment before it can be sent.

### Add-in is unavailable

If the add-in is unavailable when a message or appointment is being sent (for example, an error occurs that prevents the add-in from loading), the user is alerted. The options available to the user differ depending on the send mode option applied to the add-in.

If the **prompt user** or **soft block** option is used, the user can choose **Send Anyway** to send the item without the add-in checking it, or **Try Later** to let the item be checked by the add-in when it becomes available again.

:::image type="content" source="../images/outlook-soft-block-prompt-user-unavailable.png" alt-text="Dialog that alerts the user that the add-in is unavailable and gives the user the option to send the item now or later.":::

If the **block** option is used, the user can't send the item until the add-in becomes available.

:::image type="content" source="../images/outlook-hard-block-unavailable.png" alt-text="Dialog that alerts the user that the add-in is unavailable. The user can only send the item when the add-in is available again.":::

> [!IMPORTANT]
> If a Smart Alerts add-in that implements the [send mode override](smart-alerts-onmessagesend-walkthrough.md#override-the-send-mode-option-at-runtime-optional) feature can't complete processing an event due to an error or is unavailable when the event occurs, it uses the send mode option specified in the manifest.

### Long-running add-in operations

If the add-in runs for more than five seconds, but less than five minutes, the user is alerted that the add-in is taking longer than expected to process the message or appointment.

If the **prompt user** option is used, the user can choose **Send Anyway** to send the item without the add-in completing its check. Alternatively, the user can select **Don't Send** to stop the add-in from processing.

:::image type="content" source="../images/outlook-prompt-user-long-running.png" alt-text="Dialog that alerts the user that the add-in is taking longer than expected to process the item. The user can choose to send the item without the add-in completing its check or stop the add-in from processing the item.":::

However, if the **soft block** or **block** option is used, the user will not be able to send the item until the add-in completes processing it.

:::image type="content" source="../images/outlook-soft-hard-block-long-running.png" alt-text="Dialog that alerts the user that the add-in is taking longer than expected to process the item. The user must wait until the add-in completes processing the item before it can be sent.":::

`OnMessageSend` and `OnAppointmentSend` add-ins should be short-running and lightweight. To avoid the long-running operation dialog, use other events to process conditional checks before the `OnMessageSend` or `OnAppointmentSend` event is activated. For example, if the user is required to encrypt attachments for every message or appointment, consider using the `OnMessageAttachmentsChanged` or `OnAppointmentAttachmentsChanged` event to perform the check.

### Add-in timed out

If the add-in runs for five minutes or more, it will time out. If the **prompt user** option is used, the user can choose **Send Anyway** to send the item without the add-in completing its check. Alternatively, the user can choose **Don't Send**.

:::image type="content" source="../images/outlook-prompt-user-timeout.png" alt-text="Dialog that alerts the user that the add-in process has timed out. The user can choose to send the item without the add-in completing its check, or not send the item.":::

If the **soft block** or **block** option is  used, the user can't send the item until the add-in completes its check. The user must attempt to send the item again to reactivate the add-in.

:::image type="content" source="../images/outlook-soft-hard-block-timeout.png" alt-text="Dialog that alerts the user that the add-in process has timed out. The user must attempt to send the item again to activate the add-in before they can send the message or appointment.":::

### Intermittent or no internet connection

Event-based add-ins, including Smart Alerts add-ins, require an internet connection to launch. This section describes how an add-in behaves when Outlook is launched without connectivity, when Outlook experiences intermittent connectivity, and when the Work Offline mode, if applicable to the Outlook client, is turned on.

#### Offline when Outlook launches

When Outlook launches without internet connectivity, it's unable to determine which add-ins are installed. Because of this, Smart Alerts add-ins can't activate when the `OnMessageSend` or `OnAppointmentSend` events occur. In this scenario, to ensure that all mail items are checked for compliance before they're sent, administrators can configure policies in their organization. The policy to be set varies depending on the Outlook client.

# [Web/Windows (new)](#tab/web-new-windows)

> [!NOTE]
> When the **OnSendAddinsEnabled** policy is enabled, it will apply to all installed add-ins, not just Smart Alerts add-ins. If the policy is configured and installed add-ins don't load, the user won't be able to send items when they're offline. Administrators should only enable this policy for users in their organization if they have mandatory Smart Alerts add-ins installed.

For Outlook on the web and new Outlook on Windows, configure the **OnSendAddinsEnabled** mailbox policy in Exchange Online PowerShell. This ensures that installed add-ins always run on outgoing messages while offline.

1. [Connect to Exchange Online PowerShell](/powershell/exchange/connect-to-exchange-online-powershell).
1. Create a new mailbox policy.

   ```powershell
    New-OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > Administrators can also use an existing policy.

1. Set the **OnSendAddinsEnabled** flag to `true`.

   ```powershell
    Get-OwaMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OwaMailboxPolicy -OnSendAddinsEnabled:$true
   ```

1. Assign the policy to user mailboxes.

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'} | Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

When the **OnSendAddinsEnabled** policy is enabled and a user attempts to send a message without connectivity, the message is saved to the **Drafts** folder and a dialog is shown notifying them to reconnect.

:::image type="content" source="../images/outlook-smart-alerts-policy.png" alt-text="Dialog that alerts the user to reconnect so that the add-in can check the message before it's sent.":::

# [Windows (classic)](#tab/windows)

Outlook on Windows (classic) can load and run add-ins even if the user launches Outlook while offline. To learn more about the behavior of Smart Alerts add-ins while offline, see [Intermittent connection](#intermittent-connection).

# [Mac](#tab/mac)

For Outlook on Mac, the **OnSendAddinsWaitForLoad** mailbox key must be configured on each user's machine. This key ensures that add-ins are loaded from Exchange and are available to run checks on outgoing items. As the **OnSendAddinsWaitForLoad** key is CFPreference-compatible, it can be set by any enterprise management software for Mac, such as Jamf Pro.

> [!NOTE]
> The `OnSendAddinsWaitForLoad` policy applies to all installed add-ins, not just Smart Alerts add-ins. If the policy is configured and installed add-ins don't load, the user won't be able to send items when they're offline. Administrators should only enable this policy for users in their organization if they have mandatory Smart Alerts add-ins installed.

The following table provides details about the key.

|Field|Value|
|:---|:---|
|**Domain**|com.microsoft.outlook|
|**Key**|OnSendAddinsWaitForLoad|
|**DataType**|Boolean|
|**Possible values**|`false` (default): The currently downloaded manifests of the Smart Alerts add-ins (not necessarily the latest versions) run on outgoing mail items.<br><br>`true`: After the latest manifests of the Smart Alerts add-ins are downloaded from Exchange, the add-ins run on outgoing mail items. Otherwise, the item is blocked from being sent and the **Send** button becomes unavailable.|
|**Availability**|16.27|
|**Comments**|This key creates a policy to ensure that outgoing mail items are checked for compliance before they're sent.|

---

#### Intermittent connection

> [!NOTE]
> In classic Outlook on Windows and on Mac, the behavior of a Smart Alerts add-in is different while in [Work Offline mode](https://support.microsoft.com/office/f3a1251c-6dd5-4208-aef9-7c8c9522d633). For more information, see [Outlook client in Work Offline mode](#outlook-client-in-work-offline-mode).

If Outlook was able to load any Smart Alerts add-ins that are installed, but loses connection when a send event occurs, the behavior differs depending on the send mode option implemented by the add-in.

If the **prompt user** or **soft block** option is used, the following behavior applies.

- **Send Anyway** option is selected: The message is moved to the **Outbox** folder. When a connection is reestablished, the message is automatically sent.
- **Don't Send** option is selected: The message is saved to the **Drafts** folder. This prevents the message from being automatically sent when a connection is reestablished. When Outlook is back online and the user selects **Send**, the Smart Alerts add-in is activated.

:::image type="content" source="../images/outlook-smart-alerts-intermittent-connection.png" alt-text="Dialog that alerts the user that their message can't be processed by the add-in. The user can choose Send Anyway or Don't Send.":::

If the **block** option is used, the message is saved to the **Drafts** folder and a dialog is shown to the user notifying them to reconnect. This prevents the message from being automatically sent when a connection is reestablished. When Outlook is back online and the user selects **Send**, the Smart Alerts add-in is activated.

:::image type="content" source="../images/outlook-smart-alerts-intermittent-connection-block.png" alt-text="Dialog that alerts the user that they must reconnect to complete the add-in's check before the message can be sent.":::

#### Outlook client in Work Offline mode

In Outlook on Windows (classic client starting in Version 2310 (Build 16913.10000)) and on Mac (starting in Version 16.80 (23121017)), a Smart Alerts add-in that implements the **soft block** or **block** option can only process a mail item while the Outlook client is online. If [Work Offline mode](https://support.microsoft.com/office/f3a1251c-6dd5-4208-aef9-7c8c9522d633) is turned on in the Outlook client when a mail item is sent, the item isn't saved to the **Outbox** folder and the user is alerted that they must deactivate Work Offline mode before they can attempt to send their item.

:::image type="content" source="../images/outlook-smart-alerts-offline-mode.png" alt-text="Dialog that alerts the user that their mail item can't be processed by the Smart Alerts add-in while their Outlook client is in Work Offline mode.":::

If the Smart Alerts add-in implements the **prompt user** option, it doesn't process mail items while Work Offline mode is turned on. The item is saved to the **Outbox** folder instead.

### User navigates away from current message

When a user navigates away from the message they're sending (for example, to read a message in their inbox), the behavior of a Smart Alerts add-in differs between Outlook clients. Select the tab for the Outlook client on which the add-in is running.

# [Web / Windows](#tab/web-new-windows+windows)

#### Message composed in a window

> [!NOTE]
> In classic Outlook on Windows, the behavior described in this section applies to Version 2402 (Build 17310.10000) and later.

If a message is being composed in a separate window, such as a new message, and a user navigates away from it after they select **Send**, the Smart Alerts add-in will continue to process the message in the background. If additional actions are needed before the message can be sent, the appropriate Smart Alerts dialog is shown to the user (see [Available send mode options](#available-send-mode-options)).

#### Message composed in the Reading Pane

If a reply, forward, or existing draft is being composed in the Outlook Reading Pane, and a user navigates away from it after they select **Send**, a dialog with options is shown to the user. The options available depend on the [send mode option](#available-send-mode-options) implemented by the add-in.

If the [prompt user](#prompt-user) send mode option is implemented, the following options are shown.

- **Wait**: This option opens the message being composed in a new window, so that the Smart Alerts add-in can continue to process it. If the user navigates away from the newly opened window during processing, the add-in will continue to process the message in the background (to learn more, see [Message composed in a window](#message-composed-in-a-window)). If additional actions are needed before a message can be sent, the appropriate Smart Alerts dialog is shown to the user.
- **Send Anyway**: This option terminates the add-in operation and sends the message.

    > [!IMPORTANT]
    > In Outlook on the web and new Outlook on Windows, the **Send Anyway** option may not appear in the item-switch dialog even if the add-in currently running implements the **prompt user** send mode option. This is because the **Send Anyway** option is shown only if all installed Smart Alerts add-ins implement the **prompt user** send mode option. Otherwise, only the **Wait** and **Save as Draft** options are shown. This helps provide a more secure experience since multiple Smart Alerts add-ins can launch in any order (for more information on this behavior, see [Event-based activation behavior and limitations](../develop/event-based-activation.md#behavior-and-limitations)).

- **Save as Draft**: This option terminates the add-in and send operations and saves a draft of the message to the mailbox's **Drafts** folder.

**Outlook on the web and new Outlook on Windows**

:::image type="content" source="../images/outlook-item-switch-web-prompt-user.png" alt-text="The dialog shown in Outlook on the web and new Outlook on Windows when a user navigates away from a message being processed by a Smart Alerts add-in that implements the prompt user send mode option.":::

**Classic Outlook on Windows**

:::image type="content" source="../images/outlook-item-switch-win32-prompt-user.png" alt-text="The dialog shown in classic Outlook on Windows when a user navigates away from a message being processed by a Smart Alerts add-in that implements the prompt user send mode option.":::

If the [soft block](#soft-block) or [block](#block) send mode option is implemented, only the **Wait** and **Save as Draft** options are shown.

**Outlook on the web and new Outlook on Windows**

:::image type="content" source="../images/outlook-item-switch-web-block.png" alt-text="The dialog shown in Outlook on the web and new Outlook on Windows when a user navigates away from a message being processed by a Smart Alerts add-in that implements the soft block or block send mode option.":::

**Classic Outlook on Windows**

:::image type="content" source="../images/outlook-item-switch-win32-block.png" alt-text="The dialog shown in classic Outlook on Windows when a user navigates away from a message being processed by a Smart Alerts add-in that implements the soft block or block send mode option.":::

# [Mac](#tab/mac)

In Outlook on Mac, when a user navigates away from a message after selecting **Send**, the Smart Alerts add-in will continue to process the item in the background. If the item doesn't meet the add-in's conditions, a dialog is shown to the user to alert them that additional actions may be needed before the item can be sent. Conversely, if the item meets the add-in's conditions, the item is sent once the add-in completes processing it.

---

## Activate Smart Alerts in applications that use Simple MAPI

> [!NOTE]
> This feature is currently only supported in classic Outlook on Windows starting in Version 2301 (Build 17126.20004).

Users can send mail items through certain applications that use [Simple MAPI](/previous-versions/windows/desktop/windowsmapi/simple-mapi), even if the Outlook client isn't running at the time the item is sent. When this occurs, any installed Smart Alerts add-in won't activate to check the mail item for compliance.

To ensure that outgoing items meet the conditions of your Smart Alerts add-in before they're sent, you must turn on the **Running Outlook for Simple MAPI Mail Sending** Group Policy setting on every applicable machine in your organization.

### Behavior when the setting is turned on

When the **Running Outlook for Simple MAPI Mail Sending** setting is set to **Enabled**, users are required to have their Outlook client running at the time a mail item is sent in the following scenarios.

- A file is sent as an attachment through the **Share** > **Attach a copy instead** option in Excel, Word, or PowerPoint.

  :::image type="content" source="../images/office-attach-a-copy.png" alt-text="The 'Attach a copy instead' option selected in Word.":::

- A file is sent as an attachment through the **Send to** > **Mail recipient** option in File Explorer.

  :::image type="content" source="../images/file-explorer-send-to.png" alt-text="The 'Send to mail recipient' option selected in File Explorer.":::

- A file is sent through an application that uses Simple MAPI, which opens a new message Outlook window.

If a user's Outlook client isn't running at the time the mail item is sent, a dialog is shown to notify them that they must open their client to send the item.

:::image type="content" source="../images/outlook-simple-mapi.png" alt-text="Dialog that alerts a user to open the Outlook client when sending a mail item.":::

### Behavior when the setting is turned off or not configured

When the **Running Outlook for Simple MAPI Mail Sending** setting is set to **Disabled** or **Not Configured** in your organization, any user who uses applications that implement Simple MAPI to send mail items will be able to do so without activating their Smart Alerts add-in for compliance checks.

### Configure the Group Policy setting

By default, the **Running Outlook for Simple MAPI Mail Sending** setting is set to **Not Configured**. To turn on the setting, perform the following:

1. Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).
1. Open the **Local Group Policy Editor** (**gpedit.msc**).
1. Navigate to **User Configuration\Administrative Templates\Microsoft Outlook 2016\Miscellaneous**.
1. Open the **Running Outlook for Simple MAPI Mail Sending** setting.
1. In the dialog that appears, select **Enabled**.
1. Select **OK** or **Apply** to save your change.

## Limitations

Because the `OnMessageSend` and `OnAppointmentSend` events are supported through the event-based activation feature, the same feature limitations apply to add-ins that activate as a result of these events. For a description of these limitations, see [Event-based activation behavior and limitations](../develop/event-based-activation.md#behavior-and-limitations).

In addition to these constraints, be mindful of the following:

- Only one instance each of the `OnMessageSend` and `OnAppointmentSend` event can be declared in the manifest. If you require multiple `OnMessageSend` or `OnAppointmentSend` events, you must declare each one in a separate add-in.
- The Smart Alerts dialog message must be 500 characters or less. While you can change the message and certain aspects of a button in the Smart Alerts dialog, the following can't be customized.

  - The dialog's title bar. Your add-in's name is always displayed there.
  - The font or color of the dialog message. However, you can use Markdown to format certain elements of your message. For a list of supported elements, see [Limitations to formatting the dialog message using Markdown](#limitations-to-formatting-the-dialog-message-using-markdown).
  - The icon next to the dialog message.
  - Dialogs that provide information on event processing and progress. For example, the text and options that appear in the timeout and long-running operation dialogs can't be changed.
- In Outlook on the web and in new Outlook on Windows, the `OnAppointmentSend` event only occurs when a meeting is sent from a separate window. The event doesn't occur when a meeting is sent from an embedded meeting form or when a meeting is forwarded.

### Limitations to formatting the dialog message using Markdown

> [!NOTE]
>
> - Support for Markdown in a Smart Alerts dialog was introduced in [requirement set 1.15](/javascript/api/requirement-sets/outlook/requirement-set-1.15/outlook-requirement-set-1.15). Learn more about its [supported clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support).
> - The `errorMessageMarkdown` option of the `event.completed` call is available for preview in Outlook on Mac starting in Version 16.103 (Build 25102433). To test the property, join the [Microsoft 365 Insider program](https://techcommunity.microsoft.com/kb/microsoft-365-insider-kb/join-the-microsoft-365-insider-program-on-macos/4401756) and select the **Beta Channel** option to access Office beta builds.

You can use Markdown to format the message of a Smart Alerts dialog through the `errorMessageMarkdown` option of the `event.completed` call. However, only the following elements are supported.

- Bold, italic, or bold and italic text. Both the [asterisk (*) and underscore (_) formats](https://www.markdownguide.org/basic-syntax/#emphasis) are supported.

    ```javascript
    event.completed({
      allowEvent: false,
      ...
      errorMessageMarkdown: "**Important**: Apply the appropriate sensitivity label to your message before sending."
    });
    ```

    :::image type="content" source="../images/outlook-smart-alerts-bold.png" alt-text="A sample Smart Alerts dialog with bold text.":::

- Bulleted or unordered lists. To create an item in the list, begin with a dash (`-`) or asterisk (`*`), add the content, then append `\r` to signify item completion.

    ```javascript
    event.completed({
      allowEvent: false,
      ...
      errorMessageMarkdown: "Your email doesn't meet company guidelines.\n\nFor additional assistance, contact the IT Service Desk:\n\n- Phone number: 425-555-0102\r- Email: it@contoso.com\r- Website: [Contoso IT Service Desk](https://www.contoso.com/it-service-desk)\r"
    });
    ```

    :::image type="content" source="../images/outlook-smart-alerts-list.png" alt-text="A sample Smart Alerts dialog containing a bulleted list.":::

- Numbered or ordered lists. To create an item in the list, begin with a number followed by a period, add the content, then append `\r` to signify item completion. The first item of the list must start with the number one (`1.`) and the succeeding numbers must be in numerical order.

    ```javascript
    event.completed({
      allowEvent: false,
      ...
      errorMessageMarkdown: "Help your recipients know your intentions when you send a mail item. To set the sensitivity level of an item:\n\n1. Select **File** > **Properties**.\r2. From the **Sensitivity** dropdown, select **Normal**, **Personal**, **Private**, or **Confidential**.\r3. Select **Close**.\r"
    });
    ```

    :::image type="content" source="../images/outlook-smart-alerts-numbered-list.png" alt-text="A sample Smart Alerts dialog containing a numbered list.":::

- Links. To create a link, enclose your link text in square brackets (`[]`), then enclose the HTTPS URL in parentheses (`()`). You must provide an HTTPS URL, otherwise it won't render as a link that a user can select from the dialog. The angle brackets format (`<>`) isn't supported.

    ```javascript
    event.completed({
      allowEvent: false,
      ...
      errorMessageMarkdown: "Need onsite assistance on the day of your meeting? Visit the [Contoso Facilities](https://www.contoso.com/facilities/meetings) page to learn more."
    });
    ```

    :::image type="content" source="../images/outlook-smart-alerts-link.png" alt-text="A sample Smart Alerts dialog containing a link.":::

- New lines. Use `\n\n` to create a new line.

    ```javascript
    event.completed({
      allowEvent: false,
      ...
      errorMessageMarkdown: "Add a personalized user avatar to your signature today!\n\nTo customize your signature, visit [Customize my email signature](https://www.fabrikam.com/marketing/customize-email-signature)."
    });
    ```

    :::image type="content" source="../images/outlook-smart-alerts-new-line.png" alt-text="A sample Smart Alerts dialog containing a new line in the message.":::

> [!TIP]
> To escape characters in your message, such as an asterisk, add a backslash (`\`) before the character.

## Best practices

The Smart Alerts feature ensures that all outgoing mail items are compliant with the information protection policies of an organization and helps users improve their messages through recommendations. To ensure your add-in always provides users with a smooth and efficient sending experience, observe the following guidelines.

- **Don't let your add-in further delay the send operation**. Smart Alerts add-ins must be short-running and lightweight. Avoid overloading the `OnMessageSend` and `OnAppointmentSend` event handlers with heavy validations. To prevent this, preprocess information when other events occur, such as the `OnMessageRecipientsChanged` or `OnMessageAttachmentsChanged` event. To determine which events your add-in can respond to, see the "Supported events" section of [Activate add-ins with events](../develop/event-based-activation.md#supported-events).
- **Don't implement additional dialogs**. Prevent overwhelming your users with too many dialogs. Instead, customize the text in the Smart Alerts dialog to convey information. If needed, you can also [customize the text and functionality of certain dialog buttons](smart-alerts-onmessagesend-walkthrough.md#customize-the-text-and-functionality-of-a-button-in-the-dialog-optional) to provide users with additional information and functionality through a task pane or function.
- **Enable the appropriate Group Policy settings in your organization**. To ensure that your Smart Alerts add-in activates on each mail item, including those sent using applications that implement Simple MAPI, configure the **Running Outlook for Simple MAPI Sending** setting. To learn more about this setting, see [Activate Smart Alerts in applications that use Simple MAPI](#activate-smart-alerts-in-applications-that-use-simple-mapi).

## Debug your add-in

For guidance on how to troubleshoot your Smart Alerts add-in, see [Troubleshoot event-based and spam-reporting add-ins](../testing/troubleshoot-event-based-and-spam-reporting-add-ins.md).

## Deploy to users

For guidance on how to deploy your Smart Alerts add-in, see the "Deploy to users" section in [Activate add-ins with events](../develop/event-based-activation.md#deploy-your-add-in).

> [!IMPORTANT]
> Add-ins that use the Smart Alerts feature can only be published to Microsoft Marketplace if the manifest's [send mode property](#available-send-mode-options) is set to the **soft block** or **prompt user** option. If an add-in's send mode property is set to **block**, it can only be deployed by an organization's admin as it will fail Microsoft Marketplace validation. To learn more about publishing your event-based add-in to Microsoft Marketplace, see [Microsoft Marketplace listing options for your event-based add-in](../publish/autolaunch-store-options.md).

## Differences between Smart Alerts and the on-send feature

While Smart Alerts and the [on-send feature](outlook-on-send-addins.md) provide your users the opportunity to improve their messages and meeting invites before they're sent, Smart Alerts is a newer feature that offers you more flexibility with how you prompt your users for further action. Key differences between the two features are outlined in the following table.

|Attribute|Smart Alerts|On-send|
|-----|-----|-----|
|**Minimum supported requirement set**|[Mailbox 1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)|[Mailbox 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8)|
|**Supported Outlook clients**|<ul><li>Windows (new and classic)</li><li>Web browser (modern UI)</li><li>Mac (new UI)</li></ul>|<ul><li>Windows (new and classic)</li><li>Web browser (modern and classic UI)</li><li>Mac (new and classic UI)</li></ul>|
|**Supported manifest types**|<ul><li>Unified manifest for Microsoft 365</li><li>Add-in only manifest</li></ul>|Add-in only manifest|
|**Supported events**|**Add-in only manifest**<ul><li>`OnMessageSend`</li><li>`OnAppointmentSend`</li></ul><br>**Unified manifest for Microsoft 365**<ul><li>`"messageSending"`</li><li>`"appointmentSending"`</li></ul>|`ItemSend`|
|**Manifest extension property**|**Add-in only manifest**<ul><li>`LaunchEvent`</li></ul><br>**Unified manifest for Microsoft 365**<ul><li>`"autoRunEvents"`</li></ul>|`Events`|
|**Supported send mode options**|<ul><li>prompt user</li><li>soft block</li><li>block</li></ul><br>To learn more about each option, see [Available send mode options](#available-send-mode-options).|Block|
|**Maximum number of supported events in an add-in**|One `OnMessageSend` and one `OnAppointmentSend` event.|One `ItemSend` event.|
|**Add-in deployment**|Add-in can be published to Microsoft Marketplace if its send mode property is set to the **soft block** or **prompt user** option. Otherwise, the add-in must be deployed by an organization's administrator.|Add-in can't be published to Microsoft Marketplace. It must be deployed by an organization's administrator.|
|**Additional configuration for add-in installation**|No additional configuration is needed once the manifest is uploaded to the Microsoft 365 admin center.|Depending on the organization's compliance standards and the Outlook client used, certain mailbox policies must be configured to install the add-in.|

## See also

- [Activate add-ins with events](../develop/event-based-activation.md)
- [Microsoft Marketplace listing options for your event-based add-in](../publish/autolaunch-store-options.md)
- [Office Add-ins code sample: Office Add-ins code sample: Verify the color categories of a message or appointment before it's sent using Smart Alerts](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
- [Office Add-ins code sample: Verify the sensitivity label of a message](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-verify-sensitivity-label)
