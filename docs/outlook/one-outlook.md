---
title: Develop Outlook add-ins for the new Outlook on Windows
description: Learn how to develop add-ins that are compatible with the new Outlook on Windows.
ms.date: 07/01/2025
ms.localizationpriority: medium
---

# Develop Outlook add-ins for the new Outlook on Windows

The new Outlook on Windows desktop client unifies the Windows and web codebases to create a more consistent Outlook experience for users and administrators. Its modern and simplified interface has added capabilities and aims to improve productivity, organization, and collaboration for users. More importantly, the new Outlook on Windows supports Outlook web add-ins, so that you can continue to extend Outlook's functionality.

## Impact on VSTO and COM add-ins

The new Outlook on Windows aims to unify the extensibility experience across all Outlook platforms. To provide a more reliable and stable add-in experience, VSTO and COM add-ins aren't supported in the new Outlook on Windows. To ensure your add-in continues to work in the new Outlook on Windows, you must migrate your VSTO or COM add-in to an Outlook web add-in. Migrating to an Outlook web add-in not only enables compatibility with the new Outlook on Windows, it also makes your solution available to users on other platforms, such as Outlook on Mac, on mobile, or on the web.

To help get you started on the migration process, review the following guidance.

- The differences in features and scenarios supported by VSTO and COM add-ins and Outlook web add-ins are being addressed. To determine whether your add-in scenario is fully supported in an Outlook web add-in, see [Supported scenarios in Outlook web add-ins](#supported-scenarios-in-outlook-web-add-ins).
- For guidance on how to transition your VSTO add-in to an Outlook web add-in, see [VSTO add-in developer's guide](../overview/learning-path-transition.md) and [Tutorial: Share code between both a VSTO Add-in and an Office Add-in with a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).
- If you're new to Outlook web add-ins, try out the [Outlook quick start](../quickstarts/outlook-quickstart-yo.md) to build your first add-in.
- If you're an IT administrator and would like to learn more about how to transition to Outlook web add-ins in your organization, see [Transitioning add-ins from classic Outlook to the new Outlook for Windows](/microsoft-365-apps/outlook/get-started/migrate-com-to-web-addins#transitioning-add-ins-from-classic-outlook-to-the-new-outlook-for-windows). By default, when a user switches to new Outlook on Windows for the first time, they can choose to import their Outlook settings and install existing add-ins from classic Outlook on Windows. Because COM add-ins aren't supported in the new client, web add-in counterparts are installed instead. For more information, see [Install web add-in counterparts of existing COM add-ins in new Outlook for Windows](/microsoft-365-apps/outlook/get-started/install-web-add-ins).

> [!NOTE]
> VSTO and COM add-ins are still supported in classic Outlook on Windows.

### Supported scenarios in Outlook web add-ins

The development of the Outlook JavaScript API used by Outlook web add-ins is focused on closing the gap on scenarios that are only supported by VSTO and COM add-in solutions. This way, users who transition to the Outlook web add-in can continue to have a seamless experience.

The following table identifies key Outlook scenarios and their support status in a web add-in. This table will be updated as additional scenarios are supported. Periodically check this section as you plan to migrate your VSTO or COM add-in.

> [!TIP]
> As we continue to update the table of supported scenarios, if you want to view the recent changes made, select **Edit This Document** (**pencil icon**) from the top right corner of the article, then select **History**.
>
> To learn more about Outlook add-in features that are in preview, see [Outlook add-in API preview requirement set](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview?view=outlook-js-preview&preserve-view=true).

|Scenario|Description|Support status in Outlook web add-ins|Related features and samples|
|-----|-----|-----|-----|
|Spam email reporting and education|Enable users to report unsolicited and potentially unsafe messages and learn how to identify these messages.|Supported.|<ul><li>[Implement an integrated spam-reporting add-in](spam-reporting.md)</li><li>[ReportPhishingCommandSurface extension point](/javascript/api/manifest/extensionpoint#reportphishingcommandsurface)</li><li>[Office.context.mailbox.item.getAsFileAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getasfileasync-member(1))</li></ul>|
|Online meetings|Enable users to create and join online meetings.|Supported.|<ul><li>[Create an Outlook add-in for an online-meeting provider](online-meeting.md)</li><li>[Implement shared folders and shared mailbox scenarios in an Outlook add-in](delegate-access.md)</li><li>[Office.context.mailbox.item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item)</li><li>[Office.SharedProperties](/javascript/api/outlook/office.sharedproperties)</li></ul>|
|Meeting enhancements|Provide additional services for users when they schedule meetings, such as location selection, catering services, and room lighting and temperature adjustments.|Supported.|<ul><li>[Activate add-ins with events](../develop/event-based-activation.md)</li><li>[Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](smart-alerts-onmessagesend-walkthrough.md)</li><li>[Office.context.mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member)</li></ul>|
|Online signatures|Automatically add themed signatures to messages and appointments.|Supported.|<ul><li>[Automatically update your signature when switching between Exchange accounts](onmessagefromchanged-onappointmentfromchanged-events.md)</li><li>[Implement event-based activation in Outlook mobile add-ins](mobile-event-based.md)</li><li>[Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)</li><li>[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)</li><li>[Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body#outlook-office-body-setsignatureasync-member(1))</li></ul>|
|Customer relationship management (CRM) and tracking services|Enable users to send and retrieve information from their CRM system to track communications with existing and potential customers.|Supported.|<ul><li>[Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md)</li><li>[Log appointment notes to an external application in Outlook mobile add-ins](mobile-log-appointments.md)</li><li>[Office.context.mailbox.item.body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))</li><li>[Office.context.mailbox.item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))</li><li>[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)</li></ul>|
|Content reuse|Enable users to transfer and retrieve text and other content types from partner systems.|Supported.|<ul><li>[Prepend or append content to a message or appointment body on send](append-on-send.md)</li><li>[Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body#outlook-office-body-appendonsendasync-member(1))</li><li>[Office.context.mailbox.item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))</li><li>[Office.context.mailbox.item.body.prependOnSendAsync](/javascript/api/outlook/office.body#outlook-office-body-prependonsendasync-member(1))</li></ul>|
|Mail item transformation|Enable users to transform mail items into other formats.|Supported.|<ul><li>[getAsFileAsync method](/javascript/api/outlook/office.messageread#outlook-office-messageread-getasfileasync-member(1))</li></ul>|
|Project management|Enable users to create and track project work items from partner systems.|Supported.|<ul><li>[Activate your Outlook add-in on multiple messages](item-multi-select.md)</li><li>[Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md)</li><li>[Verify the color categories applied to a new message or appointment](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)</li></ul>|
|Attachment management|Enable users to import or export attachments from partner locations.|Supported.|<ul><li>[Activate your Outlook add-in on multiple messages](item-multi-select.md)</li><li>[Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md)</li><li>[Activate add-ins with events](../develop/event-based-activation.md)</li></ul>|
|Message encryption|Enable users to encrypt and decrypt messages.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|<ul><li>[Office.context.mailbox.item.body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))</li><li>[Office.context.mailbox.item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))</li><li>[Office.context.mailbox.item.display](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-display-member) (preview)</li><li>[Office.context.mailbox.item.display.body.setAsync](/javascript/api/outlook/office.displayedbody?view=outlook-js-preview&preserve-view=true#outlook-office-displayedbody-setasync-member(1)) (preview)</li><li>[Office.context.mailbox.item.display.subject.setAsync](/javascript/api/outlook/office.displayedsubject#outlook-office-displayedsubject-setasync-member(1)) (preview)</li></ul>|
|Data loss prevention|Prevent users from forwarding mail items that contain highly sensitive information.|Supported.|<ul><li>[Automatically check for an attachment before a message is sent](smart-alerts-onmessagesend-walkthrough.md)</li><li>[Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](onmessagesend-onappointmentsend-events.md)</li><li>[Manage the sensitivity label of your message or appointment in compose mode](sensitivity-label.md)</li><li>[Verify the sensitivity label of a message](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-verify-sensitivity-label)</li><li>[Office.SensitivityLabel](/javascript/api/outlook/office.sensitivitylabelscatalog)</li><li>[Office.SensitivityLabelsCatalog](/javascript/api/outlook/office.sensitivitylabelscatalog)</li><li>[Office.SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails)</li><li>[Office.context.mailbox.item.body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))</li><li>[Office.context.mailbox.item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))</li><li>[Office.context.mailbox.item.closeAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-closeasync-member(1))</li><li>[Office.context.mailbox.item.inReplyTo](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-inreplyto-member)</li><li>[Office.context.mailbox.item.getConversationIndexAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getconversationindexasync-member(1))</li><li>[Office.context.mailbox.item.getItemClassAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getitemclassasync-member(1))</li></ul>|
|Mail item classification|Enable users to identify and classify messages that contain sensitive information.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|<ul><li>[Automatically check for an attachment before a message is sent](smart-alerts-onmessagesend-walkthrough.md)</li><li>[Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](onmessagesend-onappointmentsend-events.md)</li><li>[Manage the sensitivity label of your message or appointment in compose mode](sensitivity-label.md)</li><li>[Manage the sensitivity level of an appointment](/javascript/api/outlook/office.sensitivity)</li><li>[Office.Sensitivity](/javascript/api/outlook/office.sensitivity)</li><li>[Office.SensitivityLabel](/javascript/api/outlook/office.sensitivitylabelscatalog)</li><li>[Office.SensitivityLabelsCatalog](/javascript/api/outlook/office.sensitivitylabelscatalog)</li><li>[Office.SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails)</li></ul>|
|Data sync service|Enable bidirectional synchronization of mail items with partner systems.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|<ul><li>[Use Microsoft Graph to manage personal contacts in Outlook](/graph/outlook-contacts-concept-overview)</li></ul>|
|Proofing mail items|Provide users with real-time proofreading assistance as they compose messages.|Not currently supported.|Not available|

There are various possibilities for extending the Outlook functionality through add-ins. If your VSTO or COM add-in solution doesn't quite fit any of the scenarios in the table, [complete the survey to share your scenario](https://aka.ms/DevNewOutlook).

## Support for classic Outlook on Windows

Classic Outlook on Windows with a Microsoft 365 subscription or a retail perpetual version of Office 2016 or later will continue to support the development of new and existing Outlook web add-ins. Additionally, it will continue to receive releases of the latest Outlook add-in features.

## Test your add-in in the new Outlook on Windows

Test your Outlook web add-in in the new Outlook on Windows today! To switch to the new Outlook on Windows, you must meet the following requirements.

- Have a Microsoft 365 work or school account connected to Exchange Online. The new client doesn't currently support on-premises, hybrid, or sovereign Exchange accounts.

    > [!NOTE]
    > While you can add non-Microsoft mail accounts, such as Gmail, to the new Outlook on Windows, you can only use Outlook add-ins with a Microsoft account. For more information, see the "Supported accounts" section of the [Outlook add-ins overview](outlook-add-ins-overview.md#supported-accounts).

- Have a minimum OS installation of Windows 10 Version 1809 (Build 17763).

To help you install and set up the Outlook desktop client, see [Getting started with the new Outlook for Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627).

For guidance on how to sideload your add-in, see [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).

> [!TIP]
>
> - If you're moving from the classic Outlook client on Windows to the new Outlook client, note that the location of add-ins is different. While add-ins are accessed from the ribbon or app bar in the classic Outlook client, access to add-ins in the new Outlook client depends on whether you're reading or composing a mail item. To learn more, see [Use add-ins in Outlook](https://support.microsoft.com/office/1ee261f9-49bf-4ba6-b3e2-2ba7bcab64c8).
> - In the new Outlook on Windows, you must keep the main client window open to run add-ins (the window can be active or inactive). If the main window is minimized, the add-in will pause or stop working.

## Debug your add-in

To debug an add-in installed in the new Outlook on Windows desktop client, perform the following:

1. Close the new Outlook on Windows client if you already have it open.
1. Then, in a command prompt, run the following command to open the new Outlook on Windows client and an instance of the Microsoft Edge DevTools.

    ```command&nbsp;line
    olk.exe --devtools
    ```

    > [!TIP]
    > After running the command, the DevTools window stays open, so that you can debug an add-in's task pane as you open and close it. If you close the DevTools window, you must close your Outlook client first before running `olk.exe --devtools` again.

1. [Sideload the add-in to Outlook on the web](sideload-outlook-add-ins-for-testing.md).
1. Use Microsoft Edge DevTools to debug your add-in.

## Add-in availability when offline

When you turn on the [offline setting](https://support.microsoft.com/office/2460e4a8-16c7-47fc-b204-b1549275aac9) in the new Outlook on Windows, you can continue to access your emails and calendar if you lose internet connection. While some functionalities remain available, Outlook add-ins and the Microsoft 365 and Copilot store aren't available when offline. The following table describes the behavior of certain types of add-ins when your machine is offline or has an intermittent connection.

| Scenario | Task pane and function command add-ins | Event-based add-ins |
| ----- | ----- | ----- |
| No internet connection when Outlook is launched | Installed add-ins don't appear on the ribbon or action bar. | Because Outlook can't determine which add-ins are installed while offline, event-based add-ins can't activate when the event they handle occurs.<br><br>In this scenario, to ensure that Smart Alerts add-ins still check messages for compliance before they're sent, administrators can configure the **OnSendAddinsEnabled** mailbox policy in Exchange Online PowerShell. When configured, outgoing messages are saved to the **Drafts** folder instead of the **Outbox** folder to prevent them from being automatically sent when the machine reconnects to the internet. For more information, see the "Offline when Outlook launches" section of [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](onmessagesend-onappointmentsend-events.md#offline-when-outlook-launches). |
| A connection is established after launching Outlook while offline | Installed add-ins appear on the ribbon and action bar. | Outlook will be able to identify which event-based add-ins are installed. Installed add-ins can then activate when the event they handle occurs.<br><br>When you select **Send** on a message that was blocked by the **OnSendAddinsEnabled** mailbox policy, the Smart Alerts add-in runs to check for compliance. |
| Machine loses connection while Outlook is in use | If you lose connection while using Outlook, your installed add-ins won't run. A dialog or notification is shown to notify that you're offline. | If you lose connection when an event occurs, the behavior differs depending on the type of event-based add-in.<ul><li>**Smart Alerts add-ins**: When you select **Send**, the behavior depends on whether the add-in implements the **prompt user**, **soft block**, or **block** send mode option. To learn more, see the "Intermittent connection" section of [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](onmessagesend-onappointmentsend-events.md#intermittent-connection).</li><li>**Other event-based add-ins**: An add-in doesn't activate when the event it handles occurs.</li></ul> |
| A connection is reestablished | Installed add-ins can run operations again. | Installed add-ins resume handling events when they occur. Messages that were moved to the **Outbox** folder are sent. When messages in the **Drafts** folder are sent, Smart Alerts add-ins are activated to ensure compliance. |

## Development experience feedback

As you test your Outlook web add-in in the new Outlook on Windows, share feedback on your experience with the developer community through [GitHub](https://github.com/OfficeDev/office-js/issues/new/choose).

## See also

- [Overview of the new Outlook for Windows](/microsoft-365-apps/outlook/overview-new-outlook-windows)
- [Blog post: Add-ins in the new Outlook for Windows](https://techcommunity.microsoft.com/t5/outlook-blog/add-ins-in-the-new-outlook-for-windows/ba-p/3954388)
- [Podcast: Update on development with new Outlook for Windows](https://www.m365devpodcast.com/update-on-development-with-new-outlook-for-windows/)
- [Outlook add-ins overview](outlook-add-ins-overview.md)
- [Build your first Outlook add-in](../quickstarts/outlook-quickstart-yo.md)
- [VSTO add-in developer's guide](../overview/learning-path-transition.md)
- [Tutorial: Share code between both a VSTO Add-in and an Office Add-in with a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md)
