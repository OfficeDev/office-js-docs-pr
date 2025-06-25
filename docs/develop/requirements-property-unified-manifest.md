---
title: Specify Office Add-in requirements in the unified manifest for Microsoft 365
description: Learn how to use requirements to configure on which host and platforms an add-in can be installed and which features are available.
ms.date: 04/18/2025
ms.topic: how-to
ms.localizationpriority: medium
---

<!-- 
  This article is deliberately left out of the Office Add-ins TOC because 
  it will be moving over to the M365 doc set as soon as that is up and running. 
-->

# Specify Office Add-in requirements in the unified manifest for Microsoft 365

There are several "requirements" properties in the [unified manifest for Microsoft 365](/office/dev/add-ins/develop/unified-manifest-overview). The [extensions.requirements](#extensionsrequirements) property controls the Office applications and versions on which the add-in can be installed. Other "requirements" properties are used to selectively suppress some features of an add-in on specific Office applications or versions where those features would be unneeded or unsupported. For more information, see [Filter features](#filter-features).

## extensions.requirements

The "extensions.requirements" property specifies the scopes, form factors, and [requirement sets](/javascript/api/requirement-sets) for Microsoft 365 Add-ins. If the Microsoft 365 version doesn't support the specified requirements, then the extension won't be available for installation. Users won't see it in the Office UI for searching and installing add-ins. Some examples:

- If the "requirements.capabilities.name" property is set to "Mailbox" and the "requirements.capabilities.minVersion" to "1.10", then the add-in isn't installable on older versions of Office that don't support **Mailbox** requirement sets greater than version 1.9.
- If the "requirements.scopes" is set to "mail", then the add-in is installable only on Outlook.
- If the "requirements.formFactors" is set to only "desktop", then the add-in isn't installable on Office running on a mobile device.

You can have more than one capability object. The following example shows how to ensure that an add-in is installable only on versions of Office that support two different requirement sets and not on mobile devices.

```json
"extensions": [
    ...
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
        ],
        "formFactors": [
            "desktop"
        ]
    }
]
```

## Filter features

The "requirements" properties in descendant objects of "extensions" are used to block some features of an add-in while still allowing the add-in to be installed. The implementation of this filtering is done at the source of installation, such as [AppSource](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center) or [Microsoft 365 Admin Center](/office/dev/add-ins/publish/publish). If the version of Office doesn't support the requirements specified for the feature, then the JSON node for the feature is removed from the manifest before it is installed in the Office application.

> [!TIP]
> Don't include a capability, formFactor, or scope requirement in a descendant object of "extensions" that's *less* restrictive than the corresponding capability, formFactor, or scope requirement in the ancestor "extensions.requirements" property, if there is one. Since the add-in can't be installed on clients that don't meet the ancestor requirement, no feature filtering would occur anyway. For example, if an "extensions.requirements.capabilities" property requires **Mailbox 1.10**, there's no point in requiring **Mailbox 1.9** in any descendant objects.

> [!NOTE]
> Office Add-ins that use the unified manifest for Microsoft 365 are *directly* supported in Office on the web, in [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), and in Office on Windows connected to a Microsoft 365 subscription, Version 2304 (Build 16320.00000) or later.
>
> When the app package that contains the unified manifest is deployed in [AppSource](https://appsource.microsoft.com/) or the [Microsoft 365 Admin Center](/office/dev/add-ins/publish/publish) then an add-in only manifest is generated from the unified manifest and stored. This add-in only manifest enables the add-in to be installed on platforms that don't directly support the unified manifest, including Office on Mac, Office on mobile, subscription versions of Office on Windows earlier than 2304 (Build 16320.00000), and perpetual versions of Office on Windows.
>
> Feature filtering is less fine-grained in the add-in only manifest. As a result, on platforms that don't directly support the unified manifest, adding a "requirements" subproperty to *any* child of "extensions" is effectively the same as adding that same "requirements" subproperty to *all* the children of "extensions" with one possible exception. So, on these platforms *none* of the features that are configured in these child properties of "extensions" will be available on platform and version combinations that don't meet the specified requirements. The exception is the "extensions.alternates" property. If this property is present, the alternates feature will be filtered in or out based only on its own "requirements" subproperty (if any), not on the "requirements" subproperties of any other child properties of "extensions".

### extensions.alternates.requirements

The "extensions.alternates" property enables add-in developers to do the following:

- Maintain a version of an add-in that was built on an older extensibility platform (such as COM or VSTO add-ins) or using the add-in only manifest, in addition to the version that uses the unified manifest.
- Either hide or give preference to the version that uses the older technology.
- Specify icons that are needed to make the unified manifest version of the add-in installable on Office versions that don't directly support the unified manifest.

For more information, see [Manage both a unified manifest and an add-in only manifest version of your Office Add-in](/office/dev/add-ins/concepts/duplicate-legacy-metaos-add-ins).

Use the "requirements" subproperty of "extensions.alternates" to selectively apply the "hide" or "prefer" subproperties only when certain requirements are met.

For example, suppose that you want to hide (from the Office UI for installing add-ins) an older version of your add-in, but only in Office versions that support the **Mailbox 1.10** requirement set. You could do that with markup similar to the following:

```json
"extensions": [
    ...
    {
        ...
        "alternates": [
            ...
            {
                ...
                "hide": {
                    "storeOfficeAddin": {
                        "officeAddinId": "b5a2794d-4aa5-4023-a84b-c60a3cbd33d4",
                        "assetId": "WA999999999"
                    }
                },
                "requirements": {
                    "capabilities": [
                        {
                            "name": "Mailbox",
                            "minVersion": "1.10"
                        }
                    ]
                }
            }
        ]
    }
]
```

### extensions.autoRunEvents.requirements

The "extensions.autoRunEvents" property configures an add-in to run specified code automatically in response to specified events. The "requirements" subproperty can be used to block this behavior in some versions of Office.

For example, suppose an Outlook add-in is configured to autolaunch in response to the **OnMailSend** event and suppose that the code in the function that runs requires the **Mailbox 1.13** requirement set. But the add-in has other features that would be useful in Office versions that only support version 1.12. To ensure that the add-in is installable in versions that support 1.12, a developer can set the "extensions.requirements.capabilities" property to the requirement set **Mailbox 1.12** instead of 1.13. But to block the autolaunch feature in versions that don't support 1.13, the developer can add an "extensions.autoRunEvents.requirements.capabilities" property that specifies **Mailbox 1.13**. The following is an example.

```json
"extensions": [
    ...
    {
        ...
        "autoRunEvents": [
            ...
            {
                ...
                "events": {
                    "type": "OnMailSend",
                    "actionId": "logOutgoingEmail",
                    "options": {
                        "sendMode": "promptUser"
                    }
                },
                "requirements": {
                    "capabilities": [
                        {
                            "name": "Mailbox",
                            "minVersion": "1.13"
                        }
                    ]
                }
            }
        ]
    }
]
```

### extensions.contentRuntimes.requirements

The "extensions.contentRuntimes" property can't be combined with any other child property of "extensions" (except "extensions.requirements"). This means that the content is the *only* feature of the add-in, so it makes no sense to filter out the feature's availability on some combinations of platform and Office version while allowing add-in to be installable on those same combinations. Accordingly, don't use the "requirements" property in "contentRuntimes". To control the installability of the content add-in, use the "extensions.requirements" property of the parent "extensions".

### extensions.contextMenus.requirements

The "extensions.contextMenus" property configures the add-in's context menus. A context menu is a shortcut menu that appears when you right-click (or select and hold) in the Office UI. The "requirements" subproperty can be used to allow context menus only when certain requirements are met.

For example, suppose you want to show context menus only in Excel versions that support the AddinCommands 1.1 requirement set. You could do that with markup similar to the following:

```json
"extensions": [
    ...
    {
        ...
        "contextMenus": [
            ...
            {
                // Insert details of the context menu configuration here.

                "requirements": {
                    "scopes": [
                        "workbook"
                    ],
                    "capabilities": [
                        {
                            "name": "AddinCommands",
                            "minVersion": "1.1"
                        }
                    ]
                }
            }
        ]
    }
]
```

### extensions.getStartedMessages.requirements

The objects in the `extensions.getStartedMessages` array provide information about an Office Add-in that appears in various places in Office, such as the callout that appears in Office when an Office Add-in is installed. There can be up to three objects in the array. If there's more than one, use the `extensions.getStartedMessages.requirements` property to ensure that no more than one of these objects is used in any given Office client. If `extensions.getStartedMessages` is omitted or all of the objects in the array are filtered out, the callout uses the values from the "name.short" and "description.short" manifest properties instead.

For example, suppose an Excel add-in simplifies the process of adding conditional formatting to ranges. Some of the APIs that the add-in uses were introduced with the **ExcelApi 1.17** requirement set, but the add-in still provides useful functionality that only requires the **ExcelApi 1.6** requirement set. The `extensions.getStartedMessages` array can be configured to provide one description of the add-in for Excel clients that support the requirement sets from **1.6** to **1.16**, but a different description for clients that support **1.17** and later. The following is an example. Note that in this example, if the add-in is configured to be installable on Excel clients that don't support requirement set **1.6**, then on those clients neither of these getStartedMessage objects would be used. Instead, Office would use the "name.short" and "description.short" properties.

```json
"extensions": [
    ...
    {
        ...
        "getStartedMessages": [
            {
                "title": "Contoso Excel Formatting",
                "description": "Use conditional formatting with our add-in.",
                "learnMoreUrl": "https://contoso.com/simple-conditional-formatting-details.html",
                "requirements": {
                    "capabilities": [
                        {
                            "name": "ExcelApi",
                            "minVersion": "1.6",
                            "maxVersion": "1.16"
                        }
                    ]
                }
            },
            {
                "title": "Contoso Advanced Excel Formatting",
                "description": "Use conditional formatting and dynamic formatting changes with our add-in.",
                "learnMoreUrl": "https://contoso.com/advanced-conditional-formatting-details.html",
                "requirements": {
                    "capabilities": [
                        {
                            "name": "ExcelApi",
                            "minVersion": "1.17"
                        }
                    ]
                }
            }
        ]
    }
]
```

### extensions.ribbons.requirements

The "extensions.ribbons" property is used to customize the Office application ribbon when the add-in is installed. The "requirements" subproperty can be used to prevent the customizations in some versions of Office.

For example, suppose an Outlook add-in is configured to add a custom button to the ribbon and the button runs a function that uses code introduced in the **Mailbox 1.9** requirement set. But the add-in has other features that would be useful on versions of Office that only support version 1.8. To ensure that the add-in is installable on versions that support 1.8, a developer can set the "extensions.requirements.capabilities" property to the requirement set **Mailbox 1.8** instead of 1.9. But to block the custom button from appearing on the ribbon in versions that don't support 1.9, the developer can add an "extensions.ribbons.requirements.capabilities" property that specifies **Mailbox 1.9**. The following is an example. For details of the custom ribbon configuration, see [Create add-in commands with the unified manifest for Microsoft 365](/office/dev/add-ins/develop/create-addin-commands-unified-manifest).

```json
"extensions": [
    ...
    {
        ...
        "ribbons": [
            ...
            {
                // Insert details of the ribbon configuration here.

                "requirements": {
                    "capabilities": [
                        {
                            "name": "Mailbox",
                            "minVersion": "1.9"
                        }
                    ]
                }
            }
        ]
    }
]
```

### extensions.runtimes.requirements

The "extensions.runtimes" property configures the sets of runtimes and actions that each extension point can use. For more information on its usage, see [Create add-in commands](/office/dev/add-ins/develop/create-addin-commands-unified-manifest), [Configure the runtime for a task pane](/office/dev/add-ins/develop/create-addin-commands-unified-manifest#configure-the-runtime-for-the-task-pane-command), and [Configure the runtime for the function command](/office/dev/add-ins/develop/create-addin-commands-unified-manifest#configure-the-runtime-for-the-function-command). For more information about runtimes in Office add-ins, see [Runtimes in Office Add-ins](/office/dev/add-ins/testing/runtimes).

The "requirements" subproperty can be used to prevent the runtime from being included in versions of Office or in Office applications where it wouldn't be used.

The previous example shown in [extensions.autoRunEvents.requirements](#extensionsautoruneventsrequirements) shows how to block the autolaunch feature in versions that don't support all of the code in the `logOutgoingEmail` function, which includes code that requires **Mailbox 1.13**. Suppose that in that same scenario, the "runtime" object that's configured to support the "logOutgoingEmail" action isn't configured to support any other action. In that case, the developer should block the runtime object in versions that don't support **Mailbox 1.13** since it would never be used. The following is an example. For details of the runtime configuration, see [Create add-in commands with the unified manifest for Microsoft 365](/office/dev/add-ins/develop/create-addin-commands-unified-manifest).

```json
"extensions": [
    ...
    {
        ...
        "runtimes": [
            ...
            {
                // Insert details of the runtime configuration here.

                "requirements": {
                    "capabilities": [
                        {
                            "name": "Mailbox",
                            "minVersion": "1.13"
                        }
                    ]
                }
            }
        ]
    }
]
```

Similarly, for the example in [extensions.ribbons.requirements](#extensionsribbonsrequirements), if the action linked to the custom button is the only action configured in a runtime object, then that runtime object should be blocked in the same circumstances in which the ribbon object is blocked.

### extensions.keyboardShortcuts.requirements (developer preview)

The `extensions.keyboardShortcuts` property defines custom keyboard shortcuts or key combinations to run specific actions. To learn how to create custom shortcuts, see [Add custom keyboard shortcuts to your Office Add-ins](../design/keyboard-shortcuts.md).

The "requirements" subproperty can be used to ensure that the custom shortcuts are only available on platforms that support the [SharedRuntime 1.1 API](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets). The following example shows how to configure this in your manifest.

```json
"extensions": [
    ...
    {
        ...
        "keyboardShortcuts": [
            {
                //Insert details of the keyboard shortcut configuration here.

                "requirements" : {
                    "capabilities": [
                        {
                            "name": "SharedRuntime",
                            "minVersion": "1.1"
                        }
                    ]
                }
            }
        ]
    }
]
```
