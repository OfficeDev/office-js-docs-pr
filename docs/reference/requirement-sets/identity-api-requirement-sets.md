---
title: Identity API requirement sets
description: ''
ms.date: 11/05/2019
ms.prod: non-product-specific
localization_priority: Normal
---

# Identity API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the Identity API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 or later on Windows<br>(one-time purchase) | Office on Windows<br>(connected to Office 365 subscription) |  Office on iPad<br>(connected to Office 365 subscription)  |  Office on Mac<br>(connected to Office 365 subscription)  | Office on the web  | SharePoint Online | OneDrive.com |Outlook.com & Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/A | Preview<b>*</b><br>(but deprecated soon) | N/A | Preview<b>*</b><br>(but deprecated soon) | Preview<b>*</b><br>(but deprecated soon) | Preview<b>*</b><br>(but deprecated soon)| N/A | N/A |
| IdentityAPI 1.2<b>&#8224;</b>  | N/A | Preview<b>*</b> | Coming soon | Preview<b>*</b> | Preview<b>*</b> | Preview<b>*</b>| Coming soon | Coming soon |

> **&#42;** During the preview phase, the Identity API requires Office 365 (the subscription version of Office). You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.
>
> **&#8224;** Version 1.2 is supported in Office on Windows 16.0.12215.20006 and later; and in Office on Mac 16.32.19102902 and later. There are some breaking changes between IdentityAPI 1.1 and 1.2. The 1.1 APIs will continue to work on the versions of Office that support the 1.2 APIs, but you must make the following changes to your code to ensure consistent behavior:
>
> - In every call of `getAccessTokenAsync` in which the option `{ forceConsent: false }` is *not* being passed, add the new option `{ allowConsentPrompt: true }`.
> - In every call of `getAccessTokenAsync` in which the option `{ forceAddAccount: true }` is being passed, add the new option `{ allowSignInPrompt: true }` (so you are passing both options).
>
> Older builds of Office will ignore the new options and newer builds will ignore the older options.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## IdentityAPI 1.2

The Single Sign On IdentityAPI 1.2 is the first released version of the API. For details about the API, see the [OfficeRuntime.Auth interface](/javascript/api/office/officeruntime.auth) reference topic.

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
