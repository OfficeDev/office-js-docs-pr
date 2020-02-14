---
title: Outlook add-in API requirement set 1.6
description: ''
ms.date: 12/17/2019
localization_priority: Normal
---

# Outlook add-in API requirement set 1.6

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.

## What's new in 1.6?

Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md). It added the following features.

- Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.
- Added a new API to open a new message form.
- Added the ability for the add-in to determine the account type of the user's mailbox.

### Change log

- Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected. Highlighted matches apply to contextual add-ins.
- Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to contextual add-ins.
- Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.
- Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#properties): Adds a new member to the user profile that indicates the type of the user's account.

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
