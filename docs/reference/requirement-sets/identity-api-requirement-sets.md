---
title: Identity API requirement sets
description: 'Identity API requirement set information for Office Add-ins.'
ms.date: 01/26/2021
ms.prod: non-product-specific
localization_priority: Normal
---

# Identity API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office Add-ins run across multiple versions of Office. The following table lists the Identity API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 or later on Windows<br>(one-time purchase) | Office on Windows<br>(connected to a Microsoft 365 subscription) |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | N/A | 2008 (build 13127.20000) or later | Coming soon | 16.40 or later | Microsoft SharePoint Online and OneDrive\* |

\* Currently, the requirement set is supported in Office on the web only for documents that are opened from Microsoft SharePoint Online and OneDrive.

> [!NOTE]
> Outlook: To require the Identity API set 1.3 in your add-in code, check if it's supported by calling `isSetSupported('IdentityAPI', '1.3')`. Declaring it in the Outlook add-in's manifest isn't supported. You can also determine if the API is supported by checking that it's not `undefined`. For further details, see [Using APIs from later requirement sets](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## IdentityAPI Preview

For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).

## See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
