---
title: Ribbon API requirement sets
description: 'Specifies which Office platforms and builds support the dynamic ribbon APIs.'
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
---

# Ribbon API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

The Ribbon API set supports programmatic control of when custom Add-in Commands (that is, custom ribbon buttons and menu items) are enabled and disabled.

Office Add-ins run across multiple versions of Office. The following table lists the Ribbon API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 on Windows<br>(one-time purchase) | Office 2016 or later on Windows<br>(one-time purchase)   | Office on Windows\*<br>(connected to a Microsoft 365 subscription) |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac\*<br>(connected to a Microsoft 365 subscription)  | Office on the web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | N/A | N/A | Version 2002 (Build 12527.20264) or later | 16.38 or later | N/A | February 2020 | N/A|

> **&#42;** During the preview phase, the Ribbon API is supported only on Excel and it requires Microsoft 365 subscription. You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). Please note that when a build graduates to the production semi-annual channel, support for preview features, including the Ribbon API, is turned off for that build.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Microsoft 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for a Microsoft 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Ribbon API 1.1

The Ribbon API 1.1 is the first version of the API. For details about the API, see the [Office.ribbon
](/javascript/api/office/office.ribbon) reference topic.

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
