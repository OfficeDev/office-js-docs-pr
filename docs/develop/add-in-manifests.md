---
title: Office Add-ins manifest
description: Get an overview of the Office Add-in manifest and its uses.
ms.topic: overview
ms.date: 10/10/2025
ms.localizationpriority: high
---

# Office Add-ins manifest

Every Office add-in has a manifest. There are two types of manifests:

- **Add-in only manifest**: This type of manifest can be used for production add-ins in Excel, OneNote, Outlook, PowerPoint, Project, and Word. It can't be used for an app that combines an add-in with some other kind of extension of the Microsoft 365 platform. Its format is XML.
- **Unified manifest for Microsoft 365**: This is an expanded version of the JSON-formatted manifest that has been used for years as the manifest for Teams Apps. Add-ins that use this manifest can be combined with other kinds of Apps for Microsoft 365, that is, extensions of Microsoft 365, in a single app that's installable as a unit. You can use this type of manifest for production Outlook add-ins. It's available for preview in Excel, PowerPoint, and Word add-ins.

[!INCLUDE [non-unified manifest clients note](../includes/non-unified-manifest-clients.md)]

The remainder of this article is applicable to both types of manifest.

> [!TIP]
>
> - For an overview that is specific to the add-in only manifest, see [Office Add-ins with an add-in only manifest](xml-manifest-overview.md).
> - For an overview that's specific to the unified manifest, see [Office Add-ins with the unified manifest for Microsoft 365](unified-manifest-overview.md).
> - If you have some familiarity with the add-in only manifest, the article [Compare the add-in only manifest with the unified manifest for Microsoft 365](json-manifest-overview.md) explains the unified manifest by comparing it with the add-in only manifest.

The manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.

A manifest file enables an Office Add-in to do the following:

- Describe itself by providing an ID, version, description, display name, and default locale.

- Specify the images used for branding the add-in and iconography used for [add-in commands](../design/add-in-commands.md) in the Office app ribbon.

- Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.

- Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.

- Declare permissions that the Office Add-in requires, such as reading or writing to the document.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## Hosting requirements

All image URIs, such as those used for [add-in commands](../design/add-in-commands.md), must support caching in production. The server hosting the image shouldn't return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.

All URLs to code or content files in the add-in should be **SSL-secured (HTTPS)**. [!INCLUDE [HTTPS guidance](../includes/https-guidance.md)]

## Best practices for submitting to Microsoft Marketplace

Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.

Add-ins submitted to Microsoft Marketplace must also include a support URL in the manifest. For more information, see [Validation policies for apps and add-ins submitted to Microsoft Marketplace](/legal/marketplace/certification-policies).

## Specify domains you want to open in the add-in window

When running in Office on the web or [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), your task pane can be navigated to any URL. However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the manifest file), that URL opens in a new browser window outside the add-in pane of the Office application.

To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the manifest. If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop. If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).

> [!NOTE]
> There are two exceptions to this behavior.
>
> - It applies only to the root pane of the add-in. If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in the manifest, even in desktop Office.
> - When a dialog is opened with the [displayDialogAsync](/javascript/api/office/office.ui?view=common-js&preserve-view=true#office-office-ui-displaydialogasync-member(1)) API, the URL that is passed to the method must be in the same domain as the add-in, but the dialog can then be directed to any URL regardless of whether it is listed in the manifest, even in desktop Office.

## Specify domains from which Office.js API calls are made

Your add-in can make Office.js API calls from the add-in's domain referenced in the manifest file. If you have other iframes within your add-in that need to access Office.js APIs, add the domain of that source URL to the manifest file. If an iframe with a source not listed in the manifest attempts to make an Office.js API call, then the add-in will receive a [permission denied error](../reference/javascript-api-for-office-error-codes.md).

## See also

- [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md)
- [Localization for Office Add-ins](localization.md)
