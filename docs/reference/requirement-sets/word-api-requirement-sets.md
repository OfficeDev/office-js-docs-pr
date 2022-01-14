---
title: Word JavaScript API requirement sets
description: 'Office Add-in requirement set information for Word builds.'
ms.date: 01/14/2022
ms.prod: word
ms.localizationpriority: high
---

# Word JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

## Requirement set availability

Word add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, iPad, and Mac. The following table lists the Word requirement sets, the Office client applications that support that requirement set, and the build or version numbers for those applications.

> [!NOTE]
> To use APIs in any of the numbered requirement sets, you should reference the **production** library on the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> For information about using preview APIs, see the [Word JavaScript preview APIs](word-preview-apis.md) article.

|  Requirement set  |   Office on Windows\*<br>(connected to a Microsoft 365 subscription)  |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web  |
|:-----|-----|:-----|:-----|:-----|
| [Preview](word-preview-apis.md) | Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Version 1612 (Build 7668.1000) or later| March 2017, 2.22 or later | March 2017, 15.32 or later| March 2017 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | December 2015 update, Version 1601 (Build 6568.1000) or later | January 2016, 1.18 or later | January 2016, 15.19 or later| September 2016 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Version 1509 (Build 4266.1001) or later| January 2016, 1.18 or later | January 2016, 15.19 or later| September 2016 |

> [!NOTE]
> Non-subscription versions of Office support requirement sets as follows:
>
> - Office 2019 and Office 2021 support WordApi 1.3 and earlier.
> - Office 2016 only supports the WordApi 1.1 requirement set.

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
