---
title: Open Browser Window requirement sets
description: 'Specifies which Office platforms and builds support the openBrowserWindow API.'
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
---

# Open Browser Window API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

The OpenBrowserWindow API set enables add-ins to open a browser to accomplish tasks that cannot always be done in the sandboxed webview control within the add-in itself; for example, downloading a PDF file when the webview control is provided by Microsoft Edge.

Office Add-ins run across multiple versions of Office. The following table lists the OpenBrowserWindow API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2021 or later on Windows<br>(one-time purchase) | Office on Windows<br>(connected to Microsoft 365 subscription) |  Office on iPad<br>(connected to Microsoft 365 subscription)  |  Office on Mac<br>(both subscription<br> and one-time purchase Office on Mac 2019 and later)   | Office on the web  |  Office Online Server  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | Build 16.0.14326.20454 or later | Version 1810 (Build 16.0.11001.20074) or later | 16.0.0.0 or later | 16.0.0.0 or later | N/A | N/A|

> [!NOTE]
> The OpenBrowserWindowApi requirement set is only available as follows:
>
> - Excel, PowerPoint, Word: Windows, Mac, iPad
> - Outlook: Windows, Mac

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Microsoft 365 Apps](/officeupdates/update-history-microsoft365-apps-by-date)
- [What version of Office am I using?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office client application](/officeupdates/update-history-microsoft365-apps-by-date)
- [Office Online Server overview](/officeonlineserver/office-online-server-overview)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## OpenBrowserWindowApi 1.1

The OpenBrowserWindowApi 1.1 is the first version of the API. For details about the API, see the [Office.context.ui](/javascript/api/office/office.context#office-office-context-ui-member) reference topic.

## See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office hosts and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
