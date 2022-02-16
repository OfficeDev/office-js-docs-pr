---
title: Identity API requirement sets
description: 'Identity API requirement set information for Office Add-ins.'
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Identity API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office Add-ins run across multiple versions of Office. The following table lists the Identity API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2021 or later on Windows<br>(one-time purchase) | Office on Windows<br>(connected to a Microsoft 365 subscription) |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(both subscription<br> and one-time purchase Office on Mac 2019 and later)   | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | Build 16.0.14326.20454 or later | Version 2008 (build 13127.20000) or later | Not supported | 16.40 or later | Microsoft SharePoint Online and OneDrive\* |

\* Currently, the requirement set is supported in Office on the web only for documents that are opened from Microsoft SharePoint Online and OneDrive.

## Outlook and Identity API requirement sets

[!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]

> [!NOTE]
> In an Outlook add-in using event-based activation, the [OfficeRuntime.Auth interface](/javascript/api/office-runtime/officeruntime.auth) is supported on Office on Windows version 2108 (build 14326.20258) or later. The [Office.Auth interface](/javascript/api/office/office.auth) is supported on version 2109 (build 14425.10000) or later. For more details according to your version, see the update history page for [Office 2021](/officeupdates/update-history-office-2021) or [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
