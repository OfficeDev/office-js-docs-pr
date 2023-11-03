---
title: Develop Outlook add-ins for the new Outlook on Windows (preview)
description: Learn how to develop add-ins that are compatible with the new Outlook on Windows (preview).
ms.date: 10/24/2023
ms.localizationpriority: medium
---

# Develop Outlook add-ins for the new Outlook on Windows (preview)

The new Outlook on Windows desktop client unifies the Windows and web codebases to create a more consistent Outlook experience for users and administrators. Its modern and simplified interface has added capabilities and aims to improve productivity, organization, and collaboration for users. More importantly, the new Outlook on Windows supports Outlook web add-ins, so that you can continue to extend Outlook's functionality.

## Impact on VSTO and COM add-ins

The new Outlook on Windows aims to unify the extensibility experience across all Outlook platforms. To provide a more reliable and stable add-in experience, VSTO and COM add-ins aren't supported in the new Outlook on Windows. To ensure your add-in continues to work in the new Outlook on Windows, you must migrate your VSTO or COM add-in to an Outlook web add-in. Migrating to an Outlook web add-in not only enables compatibility with the new Outlook on Windows, it also makes your solution available to users on other platforms, such as Outlook on Mac, on mobile, or on the web.

To help get you started on the migration process, review the following guidance.

- The differences in features and scenarios supported by VSTO and COM add-ins and Outlook web add-ins are being addressed. To determine whether your add-in scenario is fully supported in an Outlook web add-in, see [Supported scenarios in Outlook web add-ins](#supported-scenarios-in-outlook-web-add-ins).
- For guidance on how to transition your VSTO add-in to an Outlook web add-in, see [VSTO add-in developer's guide](../overview/learning-path-transition.md) and [Tutorial: Share code between both a VSTO Add-in and an Office Add-in with a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).
- If you're new to Outlook web add-ins, try out the [Outlook quick start](../quickstarts/outlook-quickstart.md) to build your first add-in.

> [!NOTE]
>
> VSTO and COM add-ins are still supported in classic Outlook on Windows.

### Supported scenarios in Outlook web add-ins

The development of the Outlook JavaScript API used by Outlook web add-ins is focused on closing the gap on scenarios that are only supported by VSTO and COM add-in solutions. This way, users who transition to the Outlook web add-in can continue to have a seamless experience.

The following table identifies key Outlook scenarios and their support status in a web add-in. This table will be updated as additional scenarios are supported. Periodically check this section as you plan to migrate your VSTO or COM add-in.

|Scenario|Description|Support status in Outlook web add-ins|Related features and samples|
|-----|-----|-----|-----|
|Spam email reporting and education|Enable users to report unsolicited and potentially unsafe messages and learn how to identify these messages.|Supported. Improvements are in development to further enhance the user experience.|[Implement an integrated spam-reporting add-in (preview)](spam-reporting.md)<br><br>[ReportPhishingCommandSurface extension point](/javascript/api/manifest/extensionpoint?view=outlook-js-preview&preserve-view=true#reportphishingcommandsurface-preview)<br><br>[Office.context.mailbox.item.getAsFileAsync](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-getasfileasync-member(1))|
|Online meetings|Enable users to create and join online meetings.|Supported.|[Create an Outlook add-in for an online-meeting provider](online-meeting.md)<br><br>[Enable shared folders and shared mailbox scenarios in an Outlook add-in](delegate-access.md)<br><br>[Office.context.mailbox.item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item)<br><br>[Office.SharedProperties](/javascript/api/outlook/office.sharedproperties)|
|Meeting enhancements|Provide additional services for users when they schedule meetings, such as location selection, catering services, and room lighting and temperature adjustments.|Supported.|[Configure your Outlook add-in for event-based activation](autolaunch.md)<br><br>[Use Smart Alerts and the OnMessageSend and OnAppointmentSend events in your Outlook add-in](smart-alerts-onmessagesend-walkthrough.md)<br><br>[Office.context.mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member)|
|Online signatures|Automatically add themed signatures to messages and appointments.|Supported.|[Automatically update your signature when switching between Exchange accounts](onmessagefromchanged-onappointmentfromchanged-events.md)<br><br>[Implement event-based activation in Outlook mobile add-ins](mobile-event-based.md)<br><br>[Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)<br><br>[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)<br><br>[Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body#outlook-office-body-setsignatureasync-member(1))|
|Customer relationship management (CRM) and tracking services|Enable users to send and retrieve information from their CRM system to track communications with existing and potential customers.|Supported.|[Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md)<br><br>[Log appointment notes to an external application in Outlook mobile add-ins](mobile-log-appointments.md)<br><br>[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)|
|Content reuse|Enable users to transfer and retrieve text and other content types from partner systems.|Supported.|[Prepend or append content to a message or appointment body on send](append-on-send.md)<br><br>[Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body#outlook-office-body-appendonsendasync-member(1))<br><br>[Office.context.mailbox.item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))<br><br>[Office.context.mailbox.item.body.prependOnSendAsync](/javascript/api/outlook/office.body#outlook-office-body-prependonsendasync-member(1))|
|Mail item transformation|Enable users to transform mail items into other formats.|Supported.|[getAsFileAsync method](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true)|
|Project management|Enable users to create and track project work items from partner systems.|Supported.|[Activate your Outlook add-in on multiple messages](item-multi-select.md)<br><br>[Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md)<br><br>[Verify the color categories applied to a new message or appointment](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)|
|Attachment management|Enable users to import or export attachments from partner locations.|Supported.|[Activate your Outlook add-in on multiple messages](item-multi-select.md)<br><br>[Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md)<br><br>[Configure your Outlook add-in for event-based activation](autolaunch.md)|
|Message encryption|Enable users to encrypt and decrypt messages.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|[Office.context.mailbox.item.display](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#outlook-office-messageread-display-member) (preview)<br><br>[Office.context.mailbox.item.display.body.setAsync](/javascript/api/outlook/office.displayedbody?view=outlook-js-preview&preserve-view=true#outlook-office-displayedbody-setasync-member(1)) (preview)<br><br>[Office.context.mailbox.item.display.subject.setAsync](/javascript/api/outlook/office.displayedsubject#outlook-office-displayedsubject-setasync-member(1)) (preview)|
|Data loss prevention|Prevent users from forwarding mail items that contain highly sensitive information.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|[Manage the sensitivity label of your message or appointment in compose mode](sensitivity-label.md)<br><br>[Verify the sensitivity label of a message](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-verify-sensitivity-label)<br><br>[Office.SensitivityLabel](/javascript/api/outlook/office.sensitivitylabelscatalog)<br><br>[Office.SensitivitiyLabelsCatalog](/javascript/api/outlook/office.sensitivitylabelscatalog)<br><br>[Office.SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails)|
|Mail item classification|Enable users to identify and classify messages that contain sensitive information.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|[Manage the sensitivity label of your message or appointment in compose mode](sensitivity-label.md)<br><br>[Manage the sensitivity level of an appointment (preview)](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)<br><br>[Office.Sensitivity](/javascript/api/outlook/office.sensitivity)<br><br>[Office.SensitivityLabel](/javascript/api/outlook/office.sensitivitylabelscatalog)<br><br>[Office.SensitivitiyLabelsCatalog](/javascript/api/outlook/office.sensitivitylabelscatalog)<br><br>[Office.SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails)|
|Data sync service|Enable bidirectional synchronization of mail items with partner systems.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|[Use Microsoft Graph to manage personal contacts in Outlook](/graph/outlook-contacts-concept-overview)|
|Proofing mail items|Provide users with real-time proofreading assistance as they compose messages.|Not currently supported.|Not available|

There are various possibilities for extending the Outlook functionality through add-ins. If your VSTO or COM add-in solution doesn't quite fit any of the scenarios in the table, [complete the survey to share your scenario](https://aka.ms/DevNewOutlook).

## Support for classic Outlook on Windows

The classic Outlook on Windows desktop client will continue to support the development of new and existing Outlook web add-ins. Additionally, it will continue to receive releases of the latest Outlook add-in features.

## Test your add-in in the new Outlook on Windows

Test your Outlook web add-in in the new Outlook on Windows today! To switch to the new Outlook on Windows that's in preview, you must meet the following requirements.

- Have a Microsoft 365 work or school account connected to Exchange Online. The new client doesn't currently support on-premises, hybrid, or sovereign Exchange accounts.

- Have a minimum OS installation of Windows 10 Version 1809 (Build 17763).

- Be a member of the [Microsoft 365 Insider program](https://insider.microsoft365.com/join/Windows).

To help you sign up and install the Outlook desktop client, see [Getting started with the new Outlook for Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627).

For guidance on how to sideload your add-in, see [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).

## Debug your add-in

To debug an add-in installed in the new Outlook on Windows desktop client, first [sideload the add-in to Outlook on the web](sideload-outlook-add-ins-for-testing.md). Then, follow the guidance in [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md) to use your browser's developer tools for debugging.

## Development experience feedback

As you test your Outlook web add-in in the new Outlook on Windows, share feedback on your experience with the developer community through [GitHub](https://github.com/OfficeDev/office-js/issues/new/choose).

## See also

- [Blog post: New Outlook for Windows available to all Office Insiders](https://insider.office.com/blog/new-outlook-for-windows-available-to-all-office-insiders)
- [Podcast: Update on development with new Outlook for Windows](https://www.m365devpodcast.com/update-on-development-with-new-outlook-for-windows/)
- [Outlook add-ins overview](outlook-add-ins-overview.md)
- [Build your first Outlook add-in](../quickstarts/outlook-quickstart.md)
- [VSTO add-in developer's guide](../overview/learning-path-transition.md)
- [Tutorial: Share code between both a VSTO Add-in and an Office Add-in with a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md)
