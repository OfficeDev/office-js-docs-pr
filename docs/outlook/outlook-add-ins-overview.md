---
title: Outlook add-ins overview
description: Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. 
ms.date: 10/09/2019
ms.custom: scenarios:getting-started
localization_priority: Priority
---

# Outlook add-ins overview

Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:

- The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Office 365 and Outlook.com), and mobile.
- Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.
- Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.

Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.

The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.

> [!NOTE]
> When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).

## Extension points

Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:

- Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

    **An add-in with command buttons on the ribbon**

    ![Add-in Command UI-less shape](../images/uiless-command-shape.png)

- Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).

    **A contextual add-in for a highlighted entity (an address)**

    ![Shows a contextual app in a card](../images/outlook-detected-entity-card.png)


> [!NOTE]
> [Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using a supported extension point.

## Mailbox items available to add-ins

Outlook add-ins are available on messages or appointments while composing or reading, but not other item types. Outlook does not activate add-ins if the current message item, in a compose or read form, is one of the following:

- Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.

- A delivery report or notification that has the message class IPM.Report.*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.

- A draft (does not have a sender assigned to it), or in the Outlook Drafts folder.

- A .msg or .eml file which is an attachment to another message.

- A .msg or .eml file opened from the file system.

- In a shared mailbox, in another user's mailbox, in an archive mailbox, or in a public folder.

- Using a custom form.

In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities. For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).

## Supported hosts

Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web in Office 365 and Outlook.com. Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time. Please refer to articles and API references for those features to see which hosts they may or may not be supported in.


## Get started building Outlook add-ins

To get started building Outlook add-ins, try the following.

- [Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.
- [Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.


## See also

- [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)
- [Design guidelines for Office Add-ins](../design/add-in-design.md)
- [License your Office and SharePoint Add-ins](/office/dev/store/license-your-add-ins)
- [Publish your Office Add-in](../publish/publish.md)
- [Make your solutions available in AppSource and within Office](/office/dev/store/submit-to-the-office-store)
