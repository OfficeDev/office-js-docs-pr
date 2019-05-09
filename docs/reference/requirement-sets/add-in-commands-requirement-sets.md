---
title: Add-in commands requirement sets
description: ''
ms.date: 05/08/2019
ms.prod: non-product-specific
localization_priority: Priority
---

# Add-in commands requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](/office/dev/add-ins/design/add-in-commands) and [Add-in commands for Outlook](/outlook/add-ins/add-in-commands-for-outlook).

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office host applications that support the initial release version, and the build versions or number for those applications.  

| Release   |  Office 2013 on Windows<br>(one-time purchase) | Office 2016 on Windows<br>(one-time purchase) | Office 2019 on Windows<br>(one-time purchase) | Office on Windows<br>(connected to Office 365)   |  Office for iPad<br>(connected to Office 365)  |  Office for Mac<br>(connected to Office 365)  | Office Online  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Add-in commands (initial release, no requirement set) | N/A | 16.0.4678.1000 *Supported in Outlook only* | Version 1809 (Build 10827.20150) or later |Version 1603 (Build 6769.0000) or later | N/A | 15.33 or later| January 2016 |

The add-in commands 1.1 requirement set introduces the ability to [autoopen a task pane with documents](/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

The following table lists the add-in commands 1.1 requirement set, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  |  Office 2013 on Windows<br>(one-time purchase) | Office 2016 on Windows<br>(one-time purchase) | Office 2019 on Windows<br>(one-time purchase) | Office on Windows<br>(connected to Office 365)   |  Office for iPad<br>(connected to Office 365)  |  Office for Mac<br>(connected to Office 365)  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.1  | N/A | 16.0.4678.1000 *Supported in Outlook only*  | Version 1809 (Build 10827.20150) or later | Version 1705 (Build 8121.1000) or later | N/A | 15.34 or later\*| May 2017 |

>\* The [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) method will erroneously return `false` for versions 16.9 &ndash; 16.14 (inclusive), but the requirement set *is* supported on these versions.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
