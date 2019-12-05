---
title: Develop Office Add-ins
description: An introduction to developing Office Add-ins.
ms.date: 11/15/2019
localization_priority: Priority
---

# Develop Office Add-ins

> [!TIP]
> Please review [Building Office Add-ins](../overview/office-add-ins-fundamentals.md) before reading this article.

All Office Add-ins are built upon the Office Add-ins platform and share a common framework through which certain capabilities can be implemented. This means that regardless of whether you're creating an add-in for Excel, Outlook, or another Office application, you can implement features such as dialog boxes, add-in commands, task panes, and single sign on (SSO). Likewise, for any add-in you build, you'll need to understand things like host and platform availability, Office JavaScript API programming patterns, the Office Add-in manifest file, and more.

## General concepts

Development concepts that apply to multiple types of Office Add-ins (i.e., Excel, Word, Outlook, etc.) are covered here in the **Office Add-ins guidance** > **Develop** section of the documentation. Review the information here before exploring the host-specific documentation that corresponds to the add-in you're building (for example, [Excel add-ins](../excel/index.md)).

- [Host and platform availability](../overview/office-add-in-availability.md)
- [Understanding the JavaScript API for Office](understanding-the-javascript-api-for-office.md)
- [Office Add-ins XML manifest](add-in-manifests.md)
- [Authentication and authorization in Office Add-ins](overview-authn-authz.md)

The **Office Add-ins guidance** > **Develop** > **How to** section of the documentation contains articles focused on specific development concepts or tasks. For example, you'll find information about tasks such as [automatically opening a task pane with a document](automatically-open-a-task-pane-with-a-document.md), [creating add-in commands](create-addin-commands.md), and [opening a dialog box](dialog-api-in-office-add-ins.md) in the **How to** section.

## Host-specific concepts

After you're familiar with the concepts that apply to multiple types of add-ins, explore the host-specific documentation for [Excel](../excel/index.md), [OneNote](../onenote/index.md), [Outlook](../outlook/index.md), [PowerPoint](../powerpoint/index.md), [Project](../project/index.md), [Visio](../visio/index.md), or [Word](../word/index.md). Each of these sections contains documentation specifically tailored to building add-ins for a certain Office host.

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Building Office Add-ins](../overview/office-add-ins-fundamentals.md)
- [General guidance for Office Add-ins](../overview/general-guidance.md)
- [Design Office Add-ins](../design/add-in-design.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)