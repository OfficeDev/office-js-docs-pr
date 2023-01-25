---
title: Develop Outlook add-ins for new Outlook on Windows (preview)
description: Learn how to develop Outlook add-ins that are compatible with the new Outlook on Windows client.
ms.topic: article
ms.date: 01/24/2023
ms.localizationpriority: medium
---

# Develop Outlook add-ins for new Outlook on Windows (preview)

The new Outlook on Windows desktop client unifies the Windows and web codebases to create a more consistent Outlook experience for users. Its modern and simplified interface with added capabilities aims to improve productivity, organization, and collaboration for users. More importantly, new Outlook on Windows supports Outlook add-ins, so that you can continue to extend Outlook's functionality to meet your, and your customers', needs.

## Preview the new Outlook on Windows client

The new Outlook on Windows client is in preview. To test it, you must:

- Have an Exchange-backed Microsoft 365 work or school account.

- Have a minimum OS installation of Windows 10 Version 1809 (Build 17763).

- Be a member of the [Office Insider Program](https://insider.office.com/join/windows).

To help you sign up and install the app, see [Getting started with the new Outlook for Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627).

## Supported Outlook add-in features

The process to develop Outlook add-ins for the new Outlook on Windows client remains the same, but support for some of the features you're familiar with or have implemented may vary.

The following table identifies which Outlook add-in features are already supported in the new Outlook on Windows client. This table will be updated as additional features are supported. As you develop your add-in, periodically check this section to determine whether a feature you need is already supported.

| Feature | Description | Compatibility status |
|-----|-----|-----|
| [Append-on-send](append-on-send.md) | Append content to messages or appointments on send. ||
| [Contextual add-in](contextual-outlook-add-ins.md) | Initiate mail-related tasks, such as selecting an address in a message to open a map, without leaving the message or appointment. ||
| [Custom properties](metadata-for-an-outlook-add-in.md) | Set and manage custom data on a mail item. ||
| [Delegate access support](delegate-access.md) | Enable delegate access scenarios in your Outlook add-in. | |
| [Event-based activation](autolaunch.md) | Automatically run mail tasks based on certain events, such as updating your signature when the From field changes. ||
| [Item multi-select (preview)](item-multi-select.md) | Activate an add-in and perform operations on multiple selected messages in one go. ||
| [On-send](outlook-on-send-addins.md) | Validate messages and appointments before they're sent. ||
| [Online meeting provider add-in](online-meeting.md) | Create an add-in for an online meeting provider. ||
| [Internet headers](internet-headers.md) | Set and manage custom data on a mail item that persists when the item leaves the Exchange server. ||
| [Pinnable task pane](pinnable-taskpane.md) | Pin an add-in's task pane to quickly process operations in messages. ||
| [Prepend-on-send (preview)](append-on-send.md) | Prepend content to messages or appointments on send. ||
| [Roaming settings](metadata-for-an-outlook-add-in.md) | Set and manage custom data on a user's mailbox. ||
| [Sensitivity label](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview?view=outlook-js-preview&preserve-view=true) | Get or set the sensitivity label of a message or appointment. ||
| [Shared mailbox support (preview)](delegate-access.md) | Enable shared mailbox scenarios in your Outlook add-in. | |
| [Smart Alerts](smart-alerts-onmessagesend-walkthrough.md) | Validate messages and appointments before they're sent. This is a newer version of the [on-send feature](outlook-on-send-addins.md). ||

## New Outlook on Windows limitations

As you test the new Outlook on Windows client and develop Outlook add-ins to be compatible with it, be mindful of the following limitations.

- Only Exchange-backed Microsoft 365 work or school accounts are supported. The new client doesn't currently support on-premises, hybrid, or sovereign Exchange accounts.

- VSTO and COM add-ins aren't supported. For guidance on how you can transition your VSTO add-in to an Outlook add-in, see [VSTO add-in developer's guide](../overview/learning-path-transition.md). If you're new to Outlook add-ins, follow the [Outlook quick start](../quickstarts/outlook-quickstart.md) to build your first add-in.

> [!NOTE]
>
> VSTO and COM add-ins are still supported in the classic Outlook on Windows client.

## Development experience feedback

We invite you to develop and test add-ins in the new Outlook on Windows client and welcome your feedback on your experience through GitHub (see the **Feedback** section at the end of this page).

## See also

- [Blog post: New Outlook for Windows available to all Office Insiders](https://insider.office.com/blog/new-outlook-for-windows-available-to-all-office-insiders)
- [Outlook add-ins overview](outlook-add-ins-overview.md)
- [Build your first Outlook add-in](../quickstarts/outlook-quickstart.md)
- [VSTO add-in developer's guide](../overview/learning-path-transition.md)
- [Tutorial: Share code between both a VSTO Add-in and an Office Add-in with a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md)
