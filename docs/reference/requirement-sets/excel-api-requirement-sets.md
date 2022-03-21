---
title: Excel JavaScript API requirement sets
description: 'Office Add-in requirement set information for Excel builds.'
ms.date: 01/14/2022
ms.prod: excel
ms.localizationpriority: high
---

# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

## Requirement set availability

Excel add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, Mac, and iPad. The following table lists the Excel requirement sets, the Office client applications that support each requirement set, and the build versions or number for those applications.

> [!NOTE]
> To use APIs in any of the numbered requirement sets or `ExcelApiOnline`, you should reference the **production** library on the [Office.js content delivery network (CDN)](https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> For information about using preview APIs, see the [Excel JavaScript preview APIs](excel-preview-apis.md) article.

|  Requirement set  |  Office on Windows<br>(connected to a Microsoft 365 subscription)  |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Preview](excel-preview-apis.md)  | Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)). |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | N/A | N/A | N/A | Latest (see [requirement set page](excel-api-online-requirement-set.md)) |
| [ExcelApi 1.14](excel-api-1-14-requirement-set.md) | Version 2108 (Build 14326.20508) or later | 16.53 or later | 16.52 or later | October 2021 |
| [ExcelApi 1.13](excel-api-1-13-requirement-set.md) | Version 2102 (Build 13801.20738) or later | 16.50 or later | 16.50 or later | June 2021 |
| [ExcelApi 1.12](excel-api-1-12-requirement-set.md) | Version 2008 (Build 13127.20408) or later | 16.40 or later | 16.40 or later | September 2020 |
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
> Non-subscription versions of Office support requirement sets as follows:
>
> - Office 2021 supports ExcelApi 1.14 and earlier.
> - Office 2019 supports ExcelApi 1.8 and earlier.
> - Office 2016 only supports the ExcelApi 1.1 requirement set.

## Office versions and build numbers

For more information about Office versions and build numbers, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## How to use Excel requirement sets at runtime and in the manifest

> [!NOTE]
> This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).

Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.

### Checking for requirement set support at runtime

The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

### Defining requirement set support in the manifest

You can use the [Requirements element](/javascript/api/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**. If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that do not support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.

The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support ExcelApi requirement set version 1.3 or greater.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
