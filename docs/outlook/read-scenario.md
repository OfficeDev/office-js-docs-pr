---
title: Create Outlook add-ins for read forms
description: Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook.
ms.date: 07/18/2024
ms.topic: overview
ms.localizationpriority: high
---

# Create Outlook add-ins for read forms

Read add-ins activate in the Reading Pane or read inspector of Outlook. Unlike compose add-ins (Outlook add-ins that are activated when a user creates a message or appointment), read add-ins are available when users:

- View an email message, meeting request, meeting response, or meeting cancellation.

   > [!NOTE]
   > Outlook doesn't activate add-ins in read form for certain types of messages, including items that are attachments to another message, items in the Outlook Drafts folder, or items that are encrypted or protected in other ways.

- View a meeting item in which the user is an attendee.

## Locate a read add-in in Outlook

The location of an add-in in the Message Read surface depends on your Outlook client.

- **Windows (classic)**, **Mac**: Select the add-in from the ribbon of the Reading Pane or read inspector.

  :::image type="content" source="../images/outlook-message-read-surface-desktop.png" alt-text="A read add-in is selected from the ribbon of an Outlook desktop client.":::

- **Web, [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)**: Select or open a message in a new window, then select the add-in from the message action bar. If your add-in doesn't appear in the action bar, select **Apps** to view your installed add-ins.

  :::image type="content" source="../images/outlook-message-read-surface-owa.png" alt-text="A read add-in is selected from the action bar of a message in Outlook on the web.":::

To select an add-in from the Appointment Read surface, open a meeting, then select the add-in from the ribbon.

:::image type="content" source="../images/outlook-appointment-read-surface.png" alt-text="A read add-in is selected from the ribbon of an appointment in Outlook on the web.":::

## Types of add-ins available in read mode

Read add-ins can be any combination of the following types.

- [Add-in commands](../design/add-in-commands.md)
- [Contextual Outlook add-ins](contextual-outlook-add-ins.md)
- [Event-based activation add-ins](../develop/event-based-activation.md)

## See also

- [Build your first Outlook add-in](../quickstarts/outlook-quickstart-yo.md)
- [Outlook add-in APIs](apis.md)
