---
title: Excel JavaScript API requirement sets
description: 'Office Add-in requirement set information for Excel builds'
ms.date: 11/15/2019
ms.prod: excel
localization_priority: Priority
---

# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## Requirement set availability

Excel add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, Mac, and iPad. The following table lists the Excel requirement sets, the Office host applications that support each requirement set, and the build versions or number for those applications.

> [!NOTE]
> To use APIs in any of the numbered requirement sets or `ExcelApiOnline`, you should reference the **production** library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> For information about using preview APIs, see the [Excel JavaScript preview APIs](./excel-preview-apis.md) article.

|  Requirement set  |  Office on Windows<br>(connected to Office 365 subscription)  |  Office on iPad<br>(connected to Office 365 subscription)  |  Office on Mac<br>(connected to Office 365 subscription)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Preview](excel-preview-apis.md)  | Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://products.office.com/office-insider)) |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | N/A | N/A | N/A | Latest (see [requirement set page](./excel-api-online-requirement-set.md)) |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | Version 1907 (Build 11929.20306) or later | 2.30 or later | 16.30 or later | October 2019 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | Version 1903 (Build 11425.20204) or later | 2.24 or later | 16.24 or later | May 2019 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | Version 1808 (Build 10730.20102) or later | 2.17 or later | 16.17 or later | September 2018 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | Version 1801 (Build 9001.2171) or later   | 2.9 or later  | 16.9 or later  | April 2018 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | Version 1704 (Build 8201.2001) or later   | 2.2 or later  | 15.36 or later | April 2017 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | Version 1703 (Build 8067.2070) or later   | 2.2 or later  | 15.36 or later | March 2017 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | Version 1701 (Build 7870.2024) or later   | 2.2 or later  | 15.36 or later | January 2017 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | Version 1608 (Build 7369.2055) or later   | 1.27 or later | 15.27 or later | September 2016 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | Version 1601 (Build 6741.2088) or later   | 1.21 or later | 15.22 or later | January 2016 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | Version 1509 (Build 4266.1001) or later   | 1.19 or later | 15.20 or later | January 2016 |

> [!NOTE]
> The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1 requirement set.

## Office versions and build numbers

For more information about Office versions and build numbers, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel)
- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)
