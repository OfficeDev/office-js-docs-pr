---
title: Use the Office modal dialog API in your Office Add-ins
description: Learn the basics of creating a modal dialog box in an Office Add-in.
ms.date: 04/30/2022
ms.localizationpriority: medium
---

# Use the Office modal dialog API in Office Add-ins (preview)

There's a *modal* dialog API available in preview. It shouldn't be used in production add-ins, but we encourage you to experiment with it. 

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

When the modal dialog is open, the user experiences the following behavior.

- The user can't interact with the Office document.
- The user can't interact with Office application (for example, Excel or Word).
- The user also can't interact with any other instances of the same Office application.
- The user can't interact with the add-in's task pane, if any, or custom add-in commands, if any. 
- Event handlers registered by the add-in do run, but the only events that can occur are those that don't require user interaction with the document, the Office application, or the add-in, other than the dialog itself. For example, calling `Office.context.ui.messageParent` in the dialog triggers the `DialogMessageReceived` event in the parent. 

Everything in the article [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md), including the **Advanced topics and special scenarios** and **Next steps** sections (and the articles that are linked to in those sections), applies to the modal dialog as well as the non-modal dialog except that the method that opens the modal dialog is [Office.context.ui.displayModalDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaymodaldialogasync-member(1)) instead of [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)). In all the code examples, and the add-in samples linked to in the **Samples** section, you can simply replace calls of `displayDialogAsync` with calls of `displayModalDialogAsync`. 