---
title: Develop Office Add-ins
description: An introduction to developing Office Add-ins.
ms.date: 11/15/2019
localization_priority: Priority
---

# Develop Office Add-ins

> [!TIP]
> Please review [Building Office Add-ins](../overview/office-add-ins-fundamentals.md) before reading this article.

All Office Add-ins are built upon the Office Add-ins platform and share a common framework through which certain capabilities can be implemented. This means that regardless of whether you're creating an add-in for Excel, Outlook, or another Office host, you can implement features such as dialog boxes, add-in commands, task panes, and single sign on (SSO). Likewise, regardless of the type of Office Add-in you're building, you'll need to understand things like Office JavaScript API programming patterns, the Office Add-in manifest file, localization, performance, privacy, and more.

## Common concepts

Development concepts that apply to multiple types of Office Add-ins (i.e., Excel, Word, Outlook, etc.) are covered here in the **Common guidance** > **Develop** section of the documentation. Review the information here before exploring the host-specific documentation that corresponds to the type of add-in you're building (for example, [Excel add-ins](../excel/index.md)).

- [Host and platform availability](../overview/office-add-in-availability.md)
- [Understanding the JavaScript API for Office](understanding-the-javascript-api-for-office.md)
- [Office Add-ins XML manifest](add-in-manifests.md)
- [Authentication and authorization in Office Add-ins](overview-authn-authz.md)

## Host-specific concepts

After you're familiar with the common concepts that apply to multiple types of add-ins, explore the host-specific documentation for the type of add-in you're building. Each of these sections contains documentation specifically tailored to building add-ins for a specific Office host.

- [Excel add-ins documentation](../excel/index.md)
- [OneNote add-ins documentation](../onenote/index.md)
- [Outlook add-ins documentation](../outlook/index.md)
- [PowerPoint add-ins documentation](../powerpoint/index.md)
- [Project add-ins documentation](../project/index.md)
- [Visio add-ins documentation](../visio/index.md)
- [Word add-ins documentation](../word/index.md)
