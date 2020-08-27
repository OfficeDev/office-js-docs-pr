---
title: Excel JavaScript API requirement sets
description: 'Office Add-in requirement set information for Excel builds.'
ms.date: 07/10/2020
ms.prod: excel
localization_priority: Priority
---

# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

## Requirement set availability

Excel add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, Mac, and iPad. The following table lists the Excel requirement sets, the Office client applications that support each requirement set, and the build versions or number for those applications.

> [!NOTE]
> To use APIs in any of the numbered requirement sets or `ExcelApiOnline`, you should reference the **production** library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> For information about using preview APIs, see the [Excel JavaScript preview APIs](excel-preview-apis.md) article.

|  Requirement set  |  Office on Windows<br>(connected to a Microsoft 365 subscription)  |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Preview](excel-preview-apis.md)  | Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)) |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | N/A | N/A | N/A | Latest (see [requirement set page](./excel-api-online-requirement-set.md)) |
| [ExcelApi 1.11](excel-api-1-11-requirement-set.md) | Version 2002 (Build 12527.20470) or later | 16.35 or later | 16.33 or later | May 2020 |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | Version 1907 (Build 11929.20306) or later | 16.0 or later | 16.30 or later | October 2019 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | Version 1903 (Build 11425.20204) or later | 16.0 or later | 16.24 or later | May 2019 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | Version 1808 (Build 10730.20102) or later | 16.0 or later | 16.17 or later | September 2018 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | Version 1801 (Build 9001.2171) or later   | 16.0 or later  | 16.9 or later  | April 2018 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | Version 1704 (Build 8201.2001) or later   | 15.0 or later  | 15.36 or later | April 2017 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | Version 1703 (Build 8067.2070) or later   | 15.0 or later  | 15.36 or later | March 2017 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | Version 1701 (Build 7870.2024) or later   | 15.0 or later  | 15.36 or later | January 2017 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | Version 1608 (Build 7369.2055) or later   | 15.0 or later | 15.27 or later | September 2016 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | Version 1601 (Build 6741.2088) or later   | 15.0 or later | 15.22 or later | January 2016 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | Version 1509 (Build 4266.1001) or later   | 15.0 or later | 15.20 or later | January 2016 |

> [!NOTE]
> Perpetual versions of Office support requirement sets as follows:
>
> - Office 2019 supports ExcelApi 1.8 and earlier.
> - Office 2016 only supports the ExcelApi 1.1 requirement set.

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel)
- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)
