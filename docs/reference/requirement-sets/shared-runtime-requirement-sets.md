---
title: Shared runtime requirement sets
description: 'Specifies the platforms and Office applications that support the SharedRuntime APIs.'
ms.date: 03/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Shared runtime requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime. This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage. For more information, see [Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).

The following table lists the SharedRuntime 1.1 requirement set, the Office client applications that support that requirement set, and the build or version numbers for the Office application.

| Requirement set | Office 2021 or later on Windows<br>(one-time purchase) | Office on Windows<br>(connected to a Microsoft 365 subscription) | Office on iPad<br>(connected to a Microsoft 365 subscription) | Office on Mac<br>(both subscription<br> and one-time purchase Office on Mac 2019 and later)  | Office on the web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | Build 16.0.14326.20454 or later | Version 2002 (Build 12527.20092) or later | N/A | 16.35 or later | February 2020 | N/A |

> [!IMPORTANT]
> At this time, the shared JavaScript runtime is not supported on iPad or in one-time purchase versions of Office 2019 or earlier. For additional support details, see the following sections.

## Support for version 1.1 on Excel

The SharedRuntime 1.1 requirement set is released for Excel on the web, Windows, and Mac.

## Preview support for version 1.1 on Word and PowerPoint

The following table lists additional application builds that support a preview of the shared JavaScript runtime. The preview version of the shared runtime is subject to change. It is not supported for use in production environments. To get the latest build you need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is by using a Microsoft 365 subscription. If you don't already have a Microsoft 365 subscription, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).

|Office application |Build |
|-------------------|------|
|PowerPoint on Windows |Build 16.0.13218.10000 or later |
|PowerPoint on Mac |Build Build 16.46.207.0 or later |
|PowerPoint on the web | February 2022 |
|Word on Windows |Build 16.0.13218.10000 or later |
|Word on Mac |Build 16.46.207.0 or later |

## Office versions and build numbers

To find out more about versions, build numbers, and Office Online Server, see:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
