---
title: Keyboard Shortcuts requirement sets
description: 'Keyboard Shortcuts requirement set information for Office Add-ins.'
ms.date: 02/15/2022
ms.prod: non-product-specific
localization_priority: Normal
---

# Keyboard Shortcuts requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office Add-ins run across multiple versions of Office. The following table lists the Keyboard Shortcuts requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 or later on Windows<br>(one-time purchase) | Office on Windows<br>(connected to a Microsoft 365 subscription) |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(both subscription<br> and one-time purchase Office on Mac 2019 and later)   | Office on the web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| KeyboardShortcuts 1.1  | N/A | Version: 2111 (Build 14701.10000) | N/A | 16.55 | September, 2021 |

> [!NOTE]
> The **KeyboardShortcuts 1.1** requirement set is supported only in Excel.

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## KeyboardShortcuts 1.1

For details about the APIs in this requirement set, see [Office.actions](/javascript/api/office/office.actions).

## See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
