---
title: Specify Office hosts and API requirements with the unified manifest
description: Learn how to specify in the unified manifest the Office applications and API requirements for your add-in to work as expected.
ms.topic: best-practice
ms.date: 02/12/2025
ms.localizationpriority: medium
---

# Specify Office applications and API requirements with the unified manifest

> [!NOTE]
> For information about specifying requirements with the add-in only manifest, see [Specify Office hosts and API requirements with the add-in only manifest](specify-office-hosts-and-api-requirements.md).

Your Office Add-in might depend on a specific Office application (also called an Office host) or on specific members of the Office JavaScript Library (office.js). For example, your add-in might:

- Run in a single Office application (e.g., Word or Excel), or several applications.
- Make use of Office JavaScript APIs that are only available in some versions of Office. For example, the volume-licensed perpetual version of Excel 2016 doesn't support all Excel-related APIs in the Office JavaScript library.
- Be designed for use only in a mobile form factor.

In these situations, you need to ensure that your add-in is never installed on Office applications or Office versions in which it cannot run.

There are also scenarios in which you want to control which features of your add-in are visible to users based on their Office application and Office version. Three examples are:

- Your add-in has features that are useful in both Word and PowerPoint, such as text manipulation, but it has some additional features that only make sense in PowerPoint, such as slide management features. You need to hide the PowerPoint-only features when the add-in is running in Word.
- Your add-in has a feature that requires an Office JavaScript API method that is supported in some versions of an Office application, such as Microsoft 365 subscription Excel, but isn't supported in others, such as volume-licensed perpetual Excel 2016. But your add-in has other features that require only Office JavaScript API methods that *are* supported in volume-licensed perpetual Excel 2016. In this scenario, you need the add-in to be installable on that version of Excel 2016, but the feature that requires the unsupported method should be hidden from those users.
- Your add-in has features that are supported in desktop Office, but not in mobile Office.

This article helps you understand how to ensure that your add-in works as expected and reaches the broadest audience possible.

> [!NOTE]
> For a high-level view of where Office Add-ins are currently supported, see the [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets) page.

> [!TIP]
> Many of the tasks described in this article are done for you, in whole or in part, when you create your add-in project with a tool, such as the [Yeoman generator for Office Add-ins](yeoman-generator-overview.md) or one of the Office Add-in templates in Visual Studio. In such cases, please interpret the task as meaning that you should verify that it has been done.

## Use the latest Office JavaScript API library

Your add-in should load the most current version of the Office JavaScript API library from the content delivery network (CDN). To do this, be sure you have the following `<script>` tag in the first HTML file your add-in opens. Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## Specify which Office applications can host your add-in

To specify the Office applications on which your add-in can be installed, use the [`"extensions.requirements.scopes"`](/microsoft-365/extensibility/schema/requirements-extension-element#scopes) array. Specify any subset of `"mail"`, `"workbook"`, `"document"`, and `"presentation"`. The following table shows which Office application and platform combinations correspond to these values. It also shows what kind of add-in can be installed for each scope.

| Name          | Office client applications                     | Available add-in types |
|:--------------|:-----------------------------------------------|:-----------------------|
| document      | Word on the web, Windows, Mac, iPad            | Task pane              |
| mail          | Outlook on the web, Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic), Android, iOS | Mail |
| presentation  | PowerPoint on the web, Windows, Mac, iPad      | Task pane, Content     |
| workbook      | Excel on the web, Windows, Mac, iPad           | Task pane, Content     |

> [!NOTE]
> Content add-ins have an `"extensions.contentRuntimes"` property. They can't have an [`"extensions.runtimes"`](/microsoft-365/extensibility/schema/extension-runtimes-array?view=m365-app-prev&preserve-view=true) property so they can't be combined with a Task pane or Mail add-in. For more information about Content add-ins, see [Content Office Add-ins](../design/content-add-ins.md).

For example, the following JSON specifies that the add-in can install on any release of Excel, which includes Excel on the web, Windows, and iPad, but can't be installed on any other Office application.

```json
"extensions": [
    {
        "requirements": {
            "scopes": [ "workbook" ],
        },
        ...
    }
]
```

> [!NOTE]
> Office applications are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices. You usually can't specify which platform can be used to run your add-in. For example, if you specify `"workbook"`, both Excel on the web and on Windows can be used to run your add-in. However, if you specify `"mail"`, your add-in won't run on Outlook mobile clients unless you define the [mobile extension point](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface).

## Specify which Office APIs your add-in needs

You can't explicitly specify the Office versions and builds or the platforms on which your add-in should be installable, and you wouldn't want to because you would have to revise your manifest whenever support for the add-in features that your add-in uses is extended to a new version or platform. Instead, specify in the manifest the APIs that your add-in needs. Office prevents the add-in from being installed on combinations of Office version and platform that don't support the APIs and ensures that the add-in won't appear in **My Add-ins**.

> [!IMPORTANT]
> Only use the `"requirements"` property that is a direct child of [`"extensions"`](/microsoft-365/extensibility/schema/root#extensions) to specify the API members that your add-in must have to be of any significant value at all. If your add-in uses an API for some features, but has other useful features that don't require the API, you should design the add-in so that it's installable on platform and Office version combinations that don't support the API but provides a diminished experience on those combinations. For this purpose, use `"requirements"` properties that aren't direct children of `"extensions"`. For more information, see [Design for alternate experiences](#design-for-alternate-experiences).

### Requirement sets

To simplify the process of specifying the APIs that your add-in needs, Office groups most APIs together in [requirement sets](office-versions-and-requirement-sets.md). The APIs in the [Common API Object Model](understanding-the-javascript-api-for-office.md#api-models) are grouped by the development feature that they support. For example, all the APIs connected to table bindings are in the requirement set called "TableBindings 1.1". The APIs in the [Application specific object models](understanding-the-javascript-api-for-office.md#api-models) are grouped by when they were released for use in production add-ins.

Requirement sets are versioned. For example, the APIs that support [Dialog Boxes](../develop/dialog-api-in-office-add-ins.md) are in the requirement set DialogApi 1.1. When additional APIs that enable messaging from a task pane to a dialog were released, they were grouped into DialogApi 1.2, along with all the APIs in DialogApi 1.1. *Each version of a requirement set is a superset of all earlier versions.*

Requirement set support varies by Office application, the version of the Office application, and the platform on which it is running. For example, ExcelApi 1.17 isn't supported on volume-licensed perpetual versions of Office before Office 2024 but ExcelApi 1.14 is supported back to Office 2021. You want your add-in to be installable on every combination of platform and Office version that supports the APIs that it uses, so you should always specify in the manifest the *minimum* version of each requirement set that your add-in requires. Details about how to do this are later in this article.

> [!TIP]
> For more information about requirement set versioning, see [Office requirement sets availability](office-versions-and-requirement-sets.md#office-requirement-sets-availability), and for the complete lists of requirement sets and information about the APIs in each, start with [Office Add-in requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets). The reference topics for most Office.js APIs also specify the requirement set they belong to (if any).

### extensions.requirements.capabilities property

Use the [`"requirements.capabilities"`](/microsoft-365/extensibility/schema/requirements-extension-element-capabilities) property to specify the minimum requirement sets that must be supported by the Office application to install your add-in. If the Office application or platform doesn't support the requirement sets or API members specified in the `"requirements.capabilities"` property, the add-in won't run in that application or platform, and won't display in **My Add-ins**.

> [!TIP]
> All APIs in the application-specific models are in requirement sets, but some of those in the Common API model aren't. If your add-in requires an API that isn't in any requirement set, you can implement a runtime check for the availability of the API and display a message to the add-in's users if it isn't supported. For more information, see [Check for API availability at runtime](specify-api-requirements-runtime.md).

The following code example shows how to configure an add-in that is installable in all Office application and platform combinations that support the following:

- `TableBindings` requirement set, which has a minimum version of "1.1".
- `OOXML` requirement set, which has a minimum version of "1.1".

```json
"extensions": [
    {
        "requirements": {
            "capabilities": [ 
                {
                    "name": "TableBindings",
                    "minVersion": "1.1"
                },
                {
                    "name": "OOXML",
                    "minVersion": "1.1"
                }
            ],
        },
        ...
    }
]
```

> [!TIP]
> For more information and another example of using the `"extensions.requirements"` property, see the `"extensions.requirements"` section in [Specify Office Add-in requirements in the unified manifest for Microsoft 365](requirements-property-unified-manifest.md#extensionsrequirements).

### Specify the form factors on which your add-in can be installed

For an Outlook add-in, you can specify whether the add-in should be installable on desktop (includes tablets) or mobile form factors. To configure this, use the `"extensions.requirements.formFactors"` property. The following example show how to make the Outlook add-in installable on both form factors.

```json
"extensions": [
    {
        "requirements": {
            ...
            "formFactors": [
                "desktop",
                "mobile"
            ]
        },
        ...
    }
]
```

## Design for alternate experiences

The extensibility features that the Office Add-in platform provides can be usefully divided into three kinds:

- Extensibility features that are available immediately after the add-in is installed. An example of this kind of feature is [Add-in Commands](../design/add-in-commands.md), which are custom ribbon buttons and menus.
- Extensibility features that are available only when the add-in is running and that are implemented with Office.js JavaScript APIs; for example, [Dialog Boxes](../develop/dialog-api-in-office-add-ins.md).
- Extensibility features that are available only at runtime but are implemented with a combination of Office.js JavaScript and manifest configuration. Examples of these are [Excel custom functions](../excel/custom-functions-overview.md), [single sign-on](sso-in-office-add-ins.md), and [custom contextual tabs](../design/contextual-tabs.md).

If your add-in uses a specific extensibility feature for some of its functionality but has other useful functionality that doesn't require the extensibility feature, you should design the add-in so that it's installable on platform and Office version combinations that don't support the extensibility feature. It can provide a valuable, albeit diminished, experience on those combinations.

You implement this design differently depending on how the extensibility feature is implemented:

- For features implemented entirely with JavaScript, see [Check for API availability at runtime](specify-api-requirements-runtime.md).
- For features that require you to configure the manifest, see the "Filter features" section of [Specify Office Add-in requirements in the unified manifest for Microsoft 365](requirements-property-unified-manifest.md#filter-features).

## See also

- [Office Add-ins manifest](add-in-manifests.md)
- [Office Add-in requirement sets](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Specify Office Add-in requirements in the unified manifest for Microsoft 365](requirements-property-unified-manifest.md)
