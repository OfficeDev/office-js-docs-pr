---
title: Add-in commands requirement sets
description: 'Overview of Office Add-in commands requirement sets.'
ms.date: 10/05/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Add-in commands requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](../../design/add-in-commands.md) and [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md).

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office client applications that support the initial release version, and the build versions or number for those applications.  

| Release   |  Office 2013 on Windows<br>(one-time purchase) | Office 2016 on Windows<br>(one-time purchase) | Office 2019 on Windows<br>(one-time purchase) | Office 2021 on Windows<br>(one-time purchase) | Office on Windows<br>(connected to a Microsoft 365 subscription)   |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Add-in commands (initial release, no requirement set) | N/A | 16.0.4678.1000 *Supported in Outlook only* | Version 1809 (Build 10827.20150) or later| 16.0.14326.20454 or later |Version 1603 (Build 6769.0000) or later | N/A | 15.33 or later| January 2016 |

The add-in commands **1.1** requirement set introduces the ability to [autoopen a task pane with documents](../../develop/automatically-open-a-task-pane-with-a-document.md).

The add-in commands **1.3** requirement set introduces manifest markup that enables an add-in to customize the placement of a custom tab on the Office ribbon and to insert built-in Office ribbon controls into custom control groups.

The following table lists the add-in commands requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  |  Office 2013 on Windows<br>(one-time purchase) | Office 2016 on Windows<br>(one-time purchase) | Office 2019 on Windows<br>(one-time purchase) |  Office 2021 on Windows<br>(one-time purchase) | Office on Windows<br>(connected to a Microsoft 365 subscription)   |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.3  | N/A | N/A | N/A | N/A | Not supported | N/A | Not supported | November 2020 |
| AddinCommands 1.1  | N/A | 16.0.4678.1000 *Supported in Outlook only*  | Version 1809 (Build 10827.20150) or later | 16.0.14326.20454 or later | Version 1705 (Build 8121.1000) or later | N/A | 15.34 or later\*| May 2017 |

>\* The [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#isSetSupported_name__minVersion_) method will erroneously return `false` for versions 16.9 &ndash; 16.14 (inclusive), but the requirement set *is* supported on these versions.

> [!IMPORTANT]
> AddinCommands 1.3 is in preview and is *only available in PowerPoint on the web*. We recommend that you try out the markup in test and development environments only. Do not use preview markup in a production environment or within business-critical documents.

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
