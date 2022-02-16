---
title: Dialog Origin requirement sets
description: 'Learn more about the Dialog Origin requirement sets.'
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Dialog Origin requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office Add-ins run across multiple versions of Office. The following table lists the Dialog Origin requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 on Windows<br>(one-time purchase) | Office 2016 on Windows<br>(one-time purchase) | Office 2019 or later on Windows<br>(one-time purchase) | Office on Windows<br>(subscription) |  Office on iPad<br>(subscription)  |  Office on Mac<br>(both subscription<br> and one-time purchase Office on Mac 2019 and later)  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | Build<br>15.0.5371.1000<br>or later | Build<br>16.0.5200.1000<br>or later | Build<br>TBD<br>or later | TBD | 2.52 or later | 16.52 or later | July, 2021 | Version 2108<br>(Build 10377.1000)<br>or later |

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## Dialog Origin 1.1

The Dialog Origin 1.1 is the first version of the API. It provides support for cross-domain messaging between a dialog and its parent page. For details about these APIs, see the [Office.ui](/javascript/api/office/office.ui) reference topic.

## See also

- [Use the Office dialog API in Office Add-ins](../../develop/dialog-api-in-office-add-ins.md)
- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
