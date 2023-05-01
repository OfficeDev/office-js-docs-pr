---
title: Develop Outlook add-ins for the new Outlook on Windows (preview)
description: Learn how to develop add-ins that are compatible with the new Outlook on Windows (preview).
ms.date: 02/07/2023
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

|Scenario|Description|Support status in Outlook web add-ins|
|-----|-----|-----|
|Spam email reporting and education|Enable users to report unsolicited and potentially unsafe messages and learn how to identify these messages.|Supported. Improvements are in development to further enhance the user experience.|
|Online meetings|Enable users to create and join online meetings.|Supported.|
|Meeting enhancements|Provide additional services for users when they schedule meetings, such as location selection, catering services, and room lighting and temperature adjustments.|Supported.|
|Online signatures|Automatically add themed signatures to messages and appointments.|Supported.|
|Customer relationship management (CRM) and tracking services|Enable users to send and retrieve information from their CRM system to track communications with existing and potential customers.|Supported.|
|Content reuse|Enable users to transfer and retrieve text and other content types from partner systems.|Supported.|
|Mail item transformation|Enable users to transform mail items into other formats.|Supported.|
|Project management|Enable users to create and track project work items from partner systems.|Supported.|
|Attachment management|Enable users to import or export attachments from partner locations.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|
|Message encryption|Enable users to encrypt and decrypt messages.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|
|Data loss prevention|Prevent users from forwarding mail items that contain highly sensitive information.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|
|Mail item classification|Enable users to identify and classify messages that contain sensitive information.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|
|Data sync service|Enable bidirectional synchronization of mail items with partner systems.|Partially supported. Essential features are yet to be addressed to create a similar experience to VSTO or COM add-ins.|
|Proofing mail items|Provide users with real-time proofreading assistance as they compose messages.|Not currently supported.|

There are various possibilities for extending the Outlook functionality through add-ins. If your VSTO or COM add-in solution doesn't quite fit any of the scenarios in the table, [complete the survey to share your scenario](https://aka.ms/DevNewOutlook).

## Support for classic Outlook on Windows

The classic Outlook on Windows desktop client will continue to support the development of new and existing Outlook web add-ins. Additionally, it will continue to receive releases of the latest Outlook add-in features.

## Test your add-in in the new Outlook on Windows

Test your Outlook web add-in in the new Outlook on Windows today! To switch to the new Outlook on Windows that's in preview, you must meet the following requirements.

- Have a Microsoft 365 work or school account connected to Exchange Online. The new client doesn't currently support on-premises, hybrid, or sovereign Exchange accounts.

- Have a minimum OS installation of Windows 10 Version 1809 (Build 17763).

- Be a member of the [Microsoft 365 Insider program](https://insider.microsoft365.com/join/windows).

To help you sign up and install the Outlook desktop client, see [Getting started with the new Outlook for Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627).

For guidance on how to sideload your add-in, see [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).

## Development experience feedback

As you test your Outlook web add-in in the new Outlook on Windows, share feedback on your experience with the developer community through [GitHub](https://github.com/OfficeDev/office-js/issues/new/choose).

## See also

- [Blog post: New Outlook for Windows available to all Office Insiders](https://insider.office.com/blog/new-outlook-for-windows-available-to-all-office-insiders)
- [Podcast: Update on development with new Outlook for Windows](https://www.m365devpodcast.com/update-on-development-with-new-outlook-for-windows/)
- [Outlook add-ins overview](outlook-add-ins-overview.md)
- [Build your first Outlook add-in](../quickstarts/outlook-quickstart.md)
- [VSTO add-in developer's guide](../overview/learning-path-transition.md)
- [Tutorial: Share code between both a VSTO Add-in and an Office Add-in with a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md)
