---
title: Understand the logic of API requirement configuration
description: Learn how Office processes API requirements specified in the manifest.
ms.date: 04/29/2026
ms.topic: best-practice
ms.localizationpriority: medium
---

# Understand the logic of API requirement configuration

You can limit which Office client applications and versions your add-in can be installed on. You can also prevent some features of the add-in from being available on certain client applications and versions. You do this by specifying in the manifest certain requirements that have to be met by an Office client before it can install the add-in, and before certain features are available. 

> [!IMPORTANT]
> The logic of requirements configuration in the manifest is the same whether your are limiting the installability of the add-in or the availability of its features, but you should be familiar with the difference between these two kinds of limitation tasks before you read this article. Start with [How to use the "requirements" property in the unified manifest for Microsoft 365](requirements-property-unified-manifest.md). For concision, this article uses terms like "limit the add-in/feature" to mean either limit where the add-in can be installed or limit where a feature is available. 

> [!NOTE]
> The practical examples in this article use the [unified manifest for Microsoft 365](unified-manifest-overview.md). If your add-in uses the add-in only manifest, see the section [Apply this guidance to the add-in only manifest](#apply-this-guidance-to-the-add-in-only-manifest).

There are three ways that you can limit the add-in/feature, discussed in the following sections. But see also [Limit by platform at runtime](#limit-by-platform-at-runtime).

## Limit by Office application

To limit to the add-in/feature to a proper subset of Excel, Outlook, PowerPoint, or Word, use the [`"requirements.scopes"`](/microsoft-365/extensibility/schema/requirements-extension-element#scopes-1) property. (OneNote and Project add-ins can't use the unified manifest. To work with them, see the section [Apply this guidance to the add-in only manifest](#apply-this-guidance-to-the-add-in-only-manifest).) For example, the following JSON limits the add-in to Outlook and Excel.

```json
"requirements": {
    "scopes": [ "mail", "workbook" ]
    --- Possibly other child properties of requirements here.
}
```

Keep in mind the following points about how Office interprets the `"scopes"` array.

- Include in the array the applications where you want the add-in/feature to be available. To block availability in an application, leave it out of the array.
- If you want the add-in/feature available in all applications, don't include a `"scopes"` property. It should only be used when you want to *limit* the add-in/feature to a proper subset of the applications. Including all four possible values is functionally equivalent to having no `"scopes"` property at all.
- It may be helpful to think of the array as a set of "OR" conditions. Your manifest is saying to the user's Office client, "Allow this add-in/feature if you are Outlook *or* you are Excel."

## Limit by form factor

You can limit Outlook add-ins by form factor. To limit to the add-in/feature to desktop devices (includes tablets) or to mobile devices, use the [`"requirements.formFactors"`](/microsoft-365/extensibility/schema/requirements-extension-element#formfactors) property. 

> [!NOTE]
> The possible values of the `"formFactors"` array are "desktop" and "mobile". 

For example, the following JSON limits the add-in to desktop devices.

```json
"requirements": {
    "formFactors": [ "desktop" ]
    --- Possibly other child properties of requirements here.
}
```

If the `"formFactors"` element isn't present, then the add-in/feature is available in both types of form factors. Including all possible values is functionally equivalent to having no `"formFactors"` property at all. So use the property only when you want to make it only available on just one form factor.

## Limit by requirement set support

> [!NOTE]
> This section assumes that you are familiar with the concept of [requirement sets](specify-office-hosts-and-api-requirements-unified.md#specify-which-office-apis-your-add-in-needs) in Office Add-ins.

To limit the add-in/feature to Office clients that support certain requirement sets, use the [`"requirements.capabilities"`](/microsoft-365/extensibility/schema/requirements-extension-element#capabilities) property. For example, the following JSON limits the add-in/feature to Office versions that support the **Mailbox 1.10** or later *and* **DialogApi 1.2** or later requirement sets. 

```json
"requirements": {
    "capabilities": [
        {
            "name": "Mailbox",
            "minVersion": "1.10"
        },
        {
            "name": "DialogAPI",
            "minVersion": "1.2"
        }
    ]
    --- Possibly other child properties of requirements here.
}
```

> [!NOTE]
> The `"minVersion"` property is optional. If it isn't present, Office assumes version "1.1".

Keep in mind the following points about how Office interprets the `"capabilities"` array.

- If you want the add-in/feature available in any Office client (other than those that are blocked by a scope or form factor requirement), regardless of what Office.js APIs it supports, don't include a `"capabilities"` property. It should only be used when you want to *limit* the add-in/feature to clients that support certain requirement sets.
- The Office client must support *all* of the requirement sets in the array in order for the the add-in/feature to be available.
- It may be helpful to think of the array as a set of "AND" conditions. Your manifest is saying to the user's Office client, "Allow this add-in/feature only if you support **Mailbox 1.10** *and* **DialogApi 1.2**."
- If there is no `"maxVersion"` child of the capability object, then Office interprets the `"minVersion"` as meaning "this version *or later*". 
- If there is a `"maxVersion"`, but no `"minVersion"`, then Office interprets the `"maxVersion"` as meaning "this version *or earlier*". For all requirement sets, "1.1" is the earliest version.
- If both a `"maxVersion"` and `"minVersion"` are present, then Office interprets the two properties as a unit meaning "only versions in this range (inclusive)". For example, the following JSON limits the add-in/feature to Office versions that support versions 1.6 through 1.16 of the **ExcelApi** requirement set. 

    ```json
    "requirements": {
        "capabilities": [
            {
                "name": "ExcelApi",
                "minVersion": "1.6",
                "maxVersion": "1.16"
            }
        ]
    }
    ```

## Combine types of limitations

A `"requirements"` object can include two or all three kinds of child properties. When it does, the Office add-in/feature is limited to Office clients that meet *all* of the specified requirements. For example, the following JSON ensures that the add-in/feature is available only in Outlook, only in versions that support **Mailbox 1.12** or later, and only on desktop form factors.

```json
"requirements": {
    "scopes": [ "mail" ],
    "capabilities": [
        {
            "name": "Mailbox",
            "minVersion": "1.12"
        }
    ],
    "formFactors": [ "desktop" ]
}
```

> [!NOTE]
> Adding any combination of "workbook", "document", or "presentation" to the `"scopes"` array in this example would *not* make the add-in/feature available on Excel, Word, or PowerPoint, because none of those applications support the **Mailbox** requirement set. 

## Limit installation of add-ins that support multiple Office applications

The logic of  the `"capabilities"` array can make it awkward to use requirement set limitations in add-ins that support multiple Office applications; that is, when there is more than one value in the `"scopes"` array (or there is no `"scopes"` property at all). Consider the following scenario in which there need to be limits on where the add-in can be installed, but there are no limitations for specific features in the add-in.  

The add-in should be installable in both Outlook and Excel, but in no other Office applications. So, there is a `"requirements"` child of `"extensions"` that has a `"scopes"` property like the following.

```json
"extensions": [
    {
        "requirements": {
            "scopes": [ "mail", "workbook" ]
        }
    }
    --- Other child properties of extensions here.
]
```

But the add-in's functionality in Outlook uses APIs in the **Mailbox 1.11** requirement set, while it's functionality in Excel uses APIs in the **ExcelApi 1.10** requirement set. It's natural to think that a `"capabilities"` object should be added to the requirement object like the following JSON.


```json
"extensions": [
    {
        "requirements": {
            "scopes": [ "mail", "workbook" ],
            "capabilities": [ 
                {
                    "name": "Mailbox",
                    "minVersion": "1.11"
                },
                {
                    "name": "ExcelApi",
                    "minVersion": "1.10"
                }
            ],
        }
    }
    --- Other child properties of extensions here.
]
```

But this configuration ensures that the add-in/feature won't be available on *any* version of Office because **Mailbox** is supported only in Outlook and **ExcelApi** is supported only in Excel. So, there is no Office application that supports both. 

To achieve the two goals of making the add-in installable on (1) Outlook, but available only in Outlook versions that support **Mailbox 1.11**, and (2) Excel, but available only in Excel versions that support **ExcelApi 1.10**, the "capabilities" objects must be moved out of the `"extensions.requirements"` object into one or more of the child `"extensions.{FEATURE}.requirements"` objects, where {FEATURE} is a child of `"extensions"`, such as `"runtimes"` or `"ribbons"`.

For a concrete example, let's extend the scenario to specify that the add-in implements function commands following the guidance in [Create add-in commands with the unified manifest for Microsoft 365](create-addin-commands-unified-manifest.md). Specifically, the add-in has the following characteristics.

- The add-in has custom ribbon buttons in both Outlook and Excel that trigger a function command in a **commands.js** file.
- To support the function command, the manifest has an `"extensions.runtimes.actions.actionId"` property whose value is "doSomething". 
- The `Office.onReady` method in the **commands.js** file tests the [Office.context.host](/javascript/api/office/office.context#office-office-context-host-member) and branches depending on whether it's Outlook or Excel. If it's Outlook, it calls [Office.Action.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) to link "doSomething" to a function named `doSomethingInOutlook`. If the host is Excel, it calls `associate` to link "doSomething" to a function named `doSomethingInExcel`.
- To implement the ribbon buttons, the manifest initially has a single ribbon object in a ["ribbons"](/microsoft-365/extensibility/schema/extension-ribbons-array) array. And this ribbon object has a control with an ["actionId"](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item#actionid) property set to "doSomething".

It's this last characteristic that needs to be changed to achieve the two goals. The following is the strategy. 

1. Remove the entire `"capabilities"` array from `"extensions.requirements"` object, so that it again looks as it did in the first code block of this section.
1. Copy the ribbon object, so that there are now two ribbon objects in the `"ribbons"` array.
1. In the first ribbon object, add a capability object that specifies the **Mailbox 1.11** requirement set.
1. In the second ribbon object, add a capability object that specifies the **ExcelApi 1.10** requirement set.

The `"ribbons"`array should now look similar to the following. 

```json
"ribbons": [
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.11"
                }
            ]
        },
            
        --- Other children of the Outlook ribbon object here.
        --- These might be identical to the Excel ribbon object below.
    },
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "ExcelApi",
                    "minVersion": "1.10"
                }
            ]
        },                

            
        --- Other children of the Excel ribbon object here.
        --- These might be identical to the Outlook ribbon object above.
    }
],
```

These changes have the desired effects:

- The add-in is installable on both Outlook and Excel.
- The ribbon control is only be available on Outlook versions that support **Mailbox 1.11** and on Excel versions that support **ExcelApi 1.10**.

> [!NOTE]
> The problem described in this section isn't limited to when two or more application-specific requirement sets are specified. There can be combinations of *common* requirement sets that create the same danger that the manifest configuration makes the add-in/feature unavailable in *any* Office client. For example, the **CustomXmlParts** requirement set is supported only in Word, and the **ActiveView** requirement set is supported only in PowerPoint. An add-in that's intended to be available on Word versions that support **CustomXmlParts** and on PowerPoint versions that support **ActiveView** would need to be configured using the technique described in this section: duplicating one or more child elements of `"extensions"` and giving each its own `"requirements.capabilities"`. 

## Apply this guidance to the add-in only manifest

The logic of how Office processes requirement configuration in the add-in only manifest is almost the same as it does for the unified manifest, but there are some differences. 

> [!NOTE]
> Basic guidance for limiting add-ins/features by Office applications and by requirement set support with the add-in only manifest is in [Specify Office applications and API requirements with the add-in only manifest](specify-office-hosts-and-api-requirements.md). For limiting the add-in/feature by form factor, see [Limit by form factor with the add-in only manifest](#limit-by-form-factor-with-the-add-in-only-manifest).

### Limit by Office application in the add-in only manifest

Outlook add-ins and task pane add-ins for all non-Outlook add-ins each have their own manifest schemas, so an Outlook add-in can't be combined with any other Office application in the same add-in. Content add-ins also have their own manifest schema and are supported only in Excel and PowerPoint. The remainder of this section is about limiting by application *within* the categories of task pane add-in and content add-in. 

The equivalent of the `"scopes"` property is the [Hosts](/javascript/api/manifest/hosts) element. Just as there can be a `"requirements.scopes"` as either a direct child of `"extensions"` or a child of one the other child properties of `"extensions"`, so too there can be a `<Hosts>` element in either the root of the add-in only manifest or in a child `<VersionOverrides>`. 

Having no `<Hosts>` element at all means that the add-in/feature is available in all the possible Office applications, which are Excel and PowerPoint for content add-ins and all applications except Outlook for task pane add-ins. So include a `<Hosts>` element when you want to limit the availability of the add-in/feature. 

### Limit by requirement set with the add-in only manifest

The equivalent of the `"capabilities"` property is the [Sets](/javascript/api/manifest/sets) element. It, too, can be in either the base manifest or in a child `<VersionOverrides>`.

Having no `<Sets>` element at all means that the add-in/feature is available in all the versions of Office applications regardless of what requirement sets they support. So include a `<Sets>` element when you want to limit the availability of the add-in/feature.

Just as with the unified manifest, it's possible to inadvertently create a manifest which blocks the add-in/feature for *all* Office clients. (See [Limit installation of add-ins that support multiple Office applications](#limit-installation-of-add-ins-that-support-multiple-office-applications) for a description of the problem.) For example, including [Set](/javascript/api/manifest/set) elements for both **WordApi 1.4** and **ExcelApi 1.10** would have this effect. The solution is parallel to the solution for the unified manifest. In this scenario take these steps.

1. Remove the problematic `<Set>` elements from the base manifest.
1. Copy the `<VersionOverrides>` element so there are now two of them.
1. In the first `<VersionOverrides>` have a `<Hosts>` element that specifies only "Document" and a `<Sets><Set>` that specifies only **WordApi 1.4**.
1. In the second `<VersionOverrides>` have a `<Hosts>` element that specifies only "Workbook" and a `<Sets><Set>` that specifies only **ExcelApi 1.10**.

> [!NOTE]
> Limiting a feature by requirement set is less fine-grained in the add-in only manifest than in the unified manifest. With the unified manifest you can have separate `"requirements"` properties in each child property of `"extensions"`, but in the add-in only manifest you can have just one `<Requirements>` child in the `<VersionOverrides>` element, and it applies to all features configured in that `<VersionOverrides>`. For more information, see [Specify requirements in a VersionOverrides element](specify-office-hosts-and-api-requirements.md#specify-requirements-in-a-versionoverrides-element). 

### Limit by form factor with the add-in only manifest

When the add-in only manifest is used, the desktop form factor is always supported in all add-ins, and the mobile form factor is possible only for Outlook add-ins. 

For an Outlook add-in, the presence or absence of a [MobileFormFactor](/javascript/api/manifest/mobileformfactor) element, as a child of a `<VersionOverrides><Hosts><Host>` element, determines whether the add-in is available in mobile devices. See [Host](/javascript/api/manifest/host) for more information.

## Limit by platform at runtime

You can't limit an add-in/feature by platform (Windows, web, Mac, iOS, or Android) using the techniques of this article. There's no `"platform"` property in the `"requirements"` object in the unified manifest and no `<Platforms>` element child of the `<Requirements>` element in the add-in only manifest. Although there are some requirement sets that are only supported on one platform, they aren't allowed in the `"capabilities"` property or the `<Sets>` element. For add-ins that are deployed by Microsoft 365 Administrators in the [Integrated apps portal](../publish/publish.md), there is a workaround. Design the add-in so that it checks at runtime for the platform. If it's a platform on which you don't want the add-in to support, show a message to the user that the add-in won't work on their version of Office and suggest which platform the user should switch to. There are actually two ways to do a runtime check. 

- Check the [Office.context.platform](/javascript/api/office/office.context#office-office-context-platform-member) property.
- Call the [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#method-details) method and pass the name of a platform-specific requirement set.

For guidance, see [Understanding platform-specific requirement sets](platform-specific-requirement-sets.md) and [Check for API availability at runtime](specify-api-requirements-runtime.md).

> [!IMPORTANT]
> This workaround isn't permitted in add-ins that are submitted to [Microsoft Marketplace](https://marketplace.microsoft.com). Your add-in *must* work meaningfully (not with merely a graceful failure message) on all combinations of platform and Office application that conform to the requirement set, form factor, and Office application hosts restrictions explicitly defined in the manifest. See [Commercial marketplace certification policies -1120.3](/legal/marketplace/certification-policies#11203-functionality).

## See also

- [How to use the "requirements" property in the unified manifest for Microsoft 365](requirements-property-unified-manifest.md)
- [Specify Office applications and API requirements with the unified manifest](specify-office-hosts-and-api-requirements-unified.md)
- [Specify Office applications and API requirements with the add-in only manifest](specify-office-hosts-and-api-requirements.md)
