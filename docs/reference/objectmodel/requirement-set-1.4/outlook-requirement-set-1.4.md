---
title: Outlook add-in API requirement set 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
---

# Outlook add-in API requirement set 1.4

The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.

## What's new in 1.4?

Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.

### Change log

- Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.
- Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.
- Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.

## See also

- [Outlook add-ins](../../../outlook/outlook-add-ins-overview.md)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](../../../quickstarts/outlook-quickstart.md)
- [Requirement sets and supported clients](../../requirement-sets/outlook-api-requirement-sets.md)
