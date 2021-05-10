---
title: Ribbon API requirement sets
description: 'Specifies which Office platforms and builds support the dynamic ribbon APIs.'
ms.date: 05/12/2021
ms.prod: non-product-specific
localization_priority: Normal
---

# Ribbon API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

The Ribbon API set supports programmatic control of when custom Add-in Commands (that is, custom ribbon buttons and menu items) are enabled and disabled.

Office Add-ins run across multiple versions of Office. The following table lists the Ribbon API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 on Windows<br>(one-time purchase) | Office 2016 or later on Windows<br>(one-time purchase)   | Office on Windows\*<br>(connected to a Microsoft 365 subscription) |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac\*<br>(connected to a Microsoft 365 subscription)  | Office on the web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | N/A | N/A | See support<br>section below | N/A | 16.38 | November, 2020 | N/A|
| RibbonApi 1.2  | N/A | N/A | 2102 (Build 13801.20294) | N/A | coming soon | May, 2021 | N/A|

> **&#42;** The Ribbon API is supported only on Excel and it requires a Microsoft 365 subscription.

## Support for version 1.1 on Office on Windows (subscription)

The 1.1 version of the RibbonApi requirement set is supported in the Consumer Channel version 2006 (build 13001.20498 or greater). For Office on Windows the feature is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available July 14th, 2020 or later. The minimum supported builds for each channel are as follows:  

|Channel | Version | Build|
|:-----|:-----|:-----|
|Current Channel | 2006 or greater | 20266.20266 or greater|
|Monthly Enterprise Channel | 2005 or greater | 12827.20538 or greater|
|Monthly Enterprise Channel | 2004 | 12730.20602 or greater|
|Semi-Annual Enterprise Channel | 2002 or greater | 12527.20880 or greater|

## More information

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Microsoft 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for a Microsoft 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Ribbon API 1.1

The Ribbon API 1.1 is the first version of the API. For details about the API, see the [Office.ribbon](/javascript/api/office/office.ribbon) reference topic.

## Ribbon API 1.2

The Ribbon API 1.2 adds support for contextual tabs. For more information, see [Create custom contextual tabs in Office Add-ins](../../design/contextual-tabs.md).

> [!NOTE]
> The **RibbonApi 1.2** requirement set is not yet supported in the manifest, so you shouldn't specify it in the manifest's `<Requirements>` section.

## See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
