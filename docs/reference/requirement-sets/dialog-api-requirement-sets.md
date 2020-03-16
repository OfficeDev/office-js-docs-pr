---
title: Dialog API requirement sets
description: ''
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
---

# Dialog API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 on Windows\*<br>(one-time purchase) | Office 2016 or later on Windows\*<br>(one-time purchase)   | Office on Windows<br>(connected to Office 365 subscription) |  Office on iPad<br>(connected to Office 365 subscription)  |  Office on Mac<br>(connected to Office 365 subscription)  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Build 15.0.4855.1000 or later | Build 16.0.4390.1000 or later | Version 1602 (Build 6741.0000) or later | 1.22 or later | 15.20 or later| January 2017 | Version 1608 (Build 7601.6800) or later|

>\* Users of the one-time purchase Office may not have accepted all patches and updates. If so, the DLL that Office uses to report its version in the UI may be greater than the versions listed here even if the updated DLLs needed to support DialogApi have not be installed on the user's computer. To ensure that the needed patch is installed, the user must go to the Office update list ([Office 2013 list](/officeupdates/msp-files-office-2013) or [Office 2016 list](/officeupdates/msp-files-office-2016)), search for **osfclient-x-none**, and install the listed patch.

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Dialog API 1.1

The Dialog API 1.1 is the first version of the API. For details about the API, see the [Dialog API](/javascript/api/office/office.ui) reference topic.

## See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office hosts and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
