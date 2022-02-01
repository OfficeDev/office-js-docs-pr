---
title: Outlook add-in API requirement set 1.4
description: 'Features and APIs that were introduced for Outlook Add-ins and the Office JavaScript APIs as part of Mailbox API 1.4.'
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# Outlook add-in API requirement set 1.4

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.4?

Requirement set 1.4 includes all of the features of [requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.

### Change log

- Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-displaydialogasync-member(1)): Displays a dialog box in an Office application.
- Added [Office.context.ui.messageParent](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-messageparent-member(1)): Delivers a message from the dialog box to its parent/opener page.
- Added [Dialog](/javascript/api/office/office.dialog?view=outlook-js-1.4&preserve-view=true) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-displaydialogasync-member(1)) method is called.

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
