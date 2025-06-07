---
title: Development lifecycle overview
description: Learn about the planning, developing, testing, and publishing lifecycle events.
author: lindalu-msft
ms.author: lindalu
ms.topic: overview
ms.date: 06/06/2025
ms.localizationpriority: high
---

# Office Add-in development lifecycle

All Office Add-ins are built upon the Office Add-ins platform. They share a common framework through which add-in capabilities are implemented. This means that regardless of whether you're creating an add-in for Excel, Outlook, or another Office application, you can have features such as dialog boxes, add-in commands, task panes, and single sign-on (SSO).

For any add-in you build, you need to understand the following concepts.

- Office application and platform availability
- Office JavaScript API programming patterns
- How to specify an add-in's settings and capabilities in the manifest file
- Troubleshooting your add-in
- Publishing your add-in

For the best foundation for these common features and application-specific implementations, review the documentation listed in the following table.

|                 |                |
| :---------------| :------------- |
| :::image type="icon" source="../images/i_best-practices_small.svg"::: |**Plan**</br>[Learn the best practices and system requirements for Office Add-ins.](../concepts/add-in-development-best-practices.md) |
| :::image type="icon" source="../images/i_code-blocks_small.svg"::: |**Develop**</br>[Learn the APIs and patterns to develop Office Add-ins.](../develop/develop-overview.md) |
| :::image type="icon" source="../images/i_recommended-testing_small.svg"::: |**Test and debug**</br>[Learn how to test and debug Office Add-ins.](../testing/test-debug-office-add-ins.md) |
| :::image type="icon" source="../images/i_deploy_small.svg"::: |**Publish**</br>[Learn how to deploy and publish Office Add-ins.](../publish/publish.md) |
| :::image type="icon" source="../images/i_reference_small.svg"::: |**Reference**</br>[View the reference documentation for the Office JavaScript APIs, the add-ins manifest, error code lists, and more.](../reference/javascript-api-for-office.md) |

## See also

- [Office Dev Center](https://developer.microsoft.com/office)
- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets)
