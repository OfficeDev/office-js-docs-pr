---
title: Word JavaScript API requirement sets
description: 'Office Add-in requirement set information for Word builds'
ms.date: 07/17/2019
ms.prod: word
localization_priority: Priority
---

# Word JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## Requirement set availability

Word add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, iPad, and Mac. The following table lists the Word requirement sets, the Office host applications that support that requirement set, and the build or version numbers for those applications.

> [!NOTE]
> To use APIs in any of the numbered requirement sets, you should reference the **production** library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> For information about using preview APIs, see the [Excel JavaScript preview APIs](word-preview-apis.md) article.

|  Requirement set  |   Office on Windows\*<br>(connected to Office 365 subscription)  |  Office on iPad<br>(connected to Office 365 subscription)  |  Office on Mac<br>(connected to Office 365 subscription)  | Office on the web  |
|:-----|-----|:-----|:-----|:-----|
| [Preview](word-preview-apis.md) | Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://products.office.com/office-insider)) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Version 1612 (Build 7668.1000) or later| March 2017, 2.22 or later | March 2017, 15.32 or later| March 2017 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | December 2015 update, Version 1601 (Build 6568.1000) or later | January 2016, 1.18 or later | January 2016, 15.19 or later| September 2016 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Version 1509 (Build 4266.1001) or later| January 2016, 1.18 or later | January 2016, 15.19 or later| September 2016 |

> [!NOTE]
> The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the WordApi 1.1 requirement set.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
