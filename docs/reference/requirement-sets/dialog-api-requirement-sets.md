---
title: Dialog API requirement sets
description: 'Learn more about the Dialog API requirement sets.'
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Dialog API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

| Requirement set | Office 2013 on Windows\*<br>(one-time purchase) | Office 2016 on Windows\*<br>(one-time purchase) | Office 2019 on Windows\*<br>(one-time purchase) | Office 2021 or later on Windows\*<br>(one-time purchase) | Office on Windows<br>(subscription) | Office on iPad<br>(subscription) |  Office on Mac<br>(both subscription<br> and one-time purchase Office on Mac 2019 and later) | Office on the web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2  | N/A | N/A | N/A | Build 16.0.14326.20454 or later | See support<br>section below | 2.37 or later | 16.37 or later | June 2020 | N/A |
| DialogApi 1.1  | Build 15.0.4855.1000 or later | Build 16.0.4390.1000 or later | Build 16.0.12527.20720 or later | Build 16.0.14326.20454 or later | Version 1602 (Build 6741.0000) or later | 1.22 or later | 15.20 or later | January 2017 | Version 1608 (Build 7601.6800) or later|

>\* Users of the one-time purchase Office may not have accepted all patches and updates. If so, the DLL that Office uses to report its version in the UI may be greater than the versions listed here even if the updated DLLs needed to support DialogApi have not be installed on the user's computer. To ensure that the needed patch is installed, the user must go to the Office update list ([Office 2013 list](/officeupdates/msp-files-office-2013) or [Office 2016 list](/officeupdates/msp-files-office-2016)), search for **osfclient-x-none**, and install the listed patch.

## Office on Windows (subscription) support

The DialogApi 1.2 requirement set is supported in the Consumer Channel version 2005 (build 12827.20268 or greater). For Office on Windows, the feature is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available June 9th, 2020 or later. The minimum supported builds for each channel are as follows:  

|Channel | Version | Build|
|:-----|:-----|:-----|
|Current Channel | 2005 or greater | 12827.20160 or greater|
|Monthly Enterprise Channel | 2004 or greater | 12730.20430 or greater|
|Semi-Annual Enterprise Channel | 2002 or greater | 12527.20720 or greater|

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Dialog API 1.1 and 1.2

The Dialog API 1.1 is the first version of the API. Requirement set 1.2 adds support for sending data from the parent page to the dialog box with the [Office.dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) method. For details about these APIs, see the [Dialog API](/javascript/api/office/office.ui) reference topic.

## See also

- [Use the Office dialog API in Office Add-ins](../../develop/dialog-api-in-office-add-ins.md)
- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
