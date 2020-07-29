---
title: Identity API requirement sets
description: 'Identity API requirement set information for Office Add-ins.'
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
---

# Identity API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office Add-ins run across multiple versions of Office. The following table lists the Identity API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 or later on Windows<br>(one-time purchase) | Office on Windows<br>(connected to a Microsoft 365 subscription) |  Office on iPad<br>(connected to a Microsoft 365 subscription)  |  Office on Mac<br>(connected to a Microsoft 365 subscription)  | Office on the web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | N/A | 2008 (build 13127.10000) or later | Coming soon | 16.40 or later | August, 2020*<b>&#8224;</b> |

> \* Initially, the requirement set is supported in Office on the web only for documents that are opened from SharePoint Online and OneDrive.com. Support for other documents will come to Office on the web later in 2020.
>
> **&#8224;** Add-ins that use the SSO APIs on these platforms will only work if the user's tenant administrator has granted consent to the add-in. The user cannot grant consent even to their own Azure AD profile.

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
- [Specify Office hosts and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
