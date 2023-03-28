---
title: Create Outlook add-ins for read forms
description: Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook.
ms.date: 03/19/2021
ms.topic: overview
ms.localizationpriority: high
---

# Create Outlook add-ins for read forms

Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook. Unlike compose add-ins (Outlook add-ins that are activated when a user is creating a message or appointment), read add-ins are available when users:

- View an email message, meeting request, meeting response, or meeting cancellation.

   > [!NOTE]
   > Outlook doesn't activate add-ins in read form for certain types of messages, including items that are attachments to another message, items in the Outlook Drafts folder, or items that are encrypted or protected in other ways.

- View a meeting item in which the user is an attendee.

- View a meeting item in which the user is the organizer (RTM release of Outlook 2013 and Exchange 2013 only).

   > [!NOTE]
   > Starting in the Office 2013 SP1 release, if the user is viewing a meeting item that the user has organized, only compose add-ins can activate and be available. Read add-ins are no longer available in this scenario.

In each of these read scenarios, Outlook activates add-ins when their activation conditions are fulfilled, and users can choose and open activated add-ins in the add-in bar in the Reading Pane or read inspector. The following figure shows the **Bing Maps** add-in activated and opened as the user is reading a message that contains a geographic address.

**The add-in pane showing the Bing Maps add-in in action for the selected Outlook message that contains an address**

![Bing Map mail app in Outlook.](../images/outlook-detected-entity-card.png)

## Types of add-ins available in read mode

Read add-ins can be any combination of the following types.

- [Add-in commands](../design/add-in-commands.md)
- [Contextual Outlook add-ins](contextual-outlook-add-ins.md)

## API features available to read add-ins

- For activating add-ins in read forms, see Table 1 in [Specify activation rules in a manifest](activation-rules.md#specify-activation-rules-in-a-manifest).
- [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md)
- [Extract entity strings from an Outlook item](extract-entity-strings-from-an-item.md)
- [Get attachments of an Outlook item from the server](get-attachments-of-an-outlook-item.md)

## See also

- [Write your first Outlook add-in](../quickstarts/outlook-quickstart.md)
