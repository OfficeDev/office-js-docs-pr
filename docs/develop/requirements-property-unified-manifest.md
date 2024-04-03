---
title: Specifying Office Add-in requirements in the unified manifest for Microsoft 365
description: Learn how to use requirements to configure on which host and platforms an add-in can be installed.
ms.date: 04/12/2024
ms.topic: how-to
ms.localizationpriority: medium
---

<!-- 
  This article is deliberately left out of the Office Add-ins TOC because 
  it will be moving over to the M365 doc set as soon as that is up and running. 
-->

# Specifying Office Add-in requirements in the unified manifest for Microsoft 365

The "requirements" property in the [unified manifest for Microsoft 365](unified-manifest-overview.md) controls the Office applications and versions on which the add-in can be installed. See [extensions.requirements](#extensions-requirements). 

It can also be used to selectively suppress some features of an add-in on specific Office applications or versions where those features would be unneeded or unsupported. For more information, see the following sections.

- [extensions.alternates.requirements](#extensionsalternatesrequirements)
- [extensions.autoRunEvents.requirements](#extensionsautoruneventsrequirements)
- [extensions.ribbons.requirements](#extensionsribbonsrequirements)
- [extensions.runtimes.requirements](#extensionsruntimesrequirements)

## extensions.requirements

The "extensions.requirements" property specifies the scopes, formFactors, and [requirement sets](/javascript/api/requirement-sets) for Microsoft 365 Add-ins. If the Microsoft 365 version doesn't support the specified requirements, then the extension won’t be available for installation. Users won't see it in the Office UI for searching and installing add-ins. Some examples:

- If the "requirements.capabilities.name" property is set to "Mailbox" and the "requirements.capabilities.minVersion" to "1.10", then the add-in isn't installable on older versions of Office that don't support the **Mailbox** requirement set greater than version 1.9.
- If the "requirements.scopes" is set to "mail", then the add-in is installable only on Outlook.
- If the "requirements.formFactors" is set to only "desktop", then the add-in isn't installable on Office running on a mobile device.

### extensions.alternates.requirements

The "extensions.alternates" property enables add-in developers to do the following:

- Maintain a version of an add-in that was built on an older extensibility platform (such as COM or VSTO add-ins) or using the XML manifest, in addition to the version that uses the unified manifest.
- Either hide or give preference to the version that uses the older technology.
- Specify icons that are needed to make the unified manifest version of the add-in installable on Office versions that don't directly support the unified manifest. 

[!INCLUDE [non-unified manifest clients note](../includes/non-unified-manifest-clients.md)]

For more information, see [Manage both a unified manifest and an XML manifest version of your Office Add-in](../concepts/duplicate-legacy-metaos-add-ins.md).

The "requirements" subproperty of "extensions.alternates" to selectively apply the "hide" or "prefer" subproperties only when certain requirements are met. 

For example, suppose that you want to hide (from the Office UI for installing add-ins) an older version of your add-in, but only in Office versions that support the **Mailbox 1.10** requirement set. You could do that with markup similar to the following. 

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
                    "capabilities": {
                        "name": "MailBox",
                        "minVersion": "1.10"
                    }
                }
            }
        ]
    }
]
```

### extensions.autoRunEvents.requirements

The "extensions.autoRunEvents" property configures an add-in to run specified code automatically in response to specified events. The "requirements" subproperty can be used to block this behavior in some versions of Office.

For example, suppose an Outlook add-in is configured to autolaunch in response to the **OnMailSend** event and code in the function that runs requires the **Mailbox 1.13** requirement set. But the add-in has other features that would be useful Office versions that only support version 1.12. To ensure that the add-in is installable on versions that support 1.12, a developer can set the “extensions.requirements.capabilities” property to the requirement set **MailBox 1.12**. But to block the autolaunch feature in versions that don't support 1.13, the developer can add an “extensions.autoRunEvents.requirements.capabilities” property that specifies **Mailbox 1.13**. The following is an example.

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
                    "capabilities": {
                        "name": "MailBox",
                        "minVersion": "1.13"
                    }
                }
            }
        ]
    }
]
```

### extensions.ribbons.requirements

The "extensions.ribbons" property is used to customize the Office application ribbon when the add-in is installed. The "requirements" subproperty can be used to prevent the customizations in some versions of Office.

For example, suppose an Outlook add-in is configured to add a custom button to the ribbon and the button runs a function that uses code introduced in the **Mailbox 1.9** requirement set. But the add-in has other features that would be useful on versions of Office that only support version 1.8. To ensure that the add-in is installable on versions that support 1.8, a developer can set the “extensions.requirements.capabilities” property to the requirement set **MailBox 1.8**. But to block the custom button from appearing on the ribbon in versions that don't support 1.9, the developer can add an “extensions.ribbons.requirements.capabilities” property that specifies **Mailbox 1.9**. The following is an example. For details of the custom ribbon configuration, see [Create add-in commands with the unified manifest for Microsoft 365](create-addin-commands-unified-manifest.md).

```json
"extensions": [
    ...
    {
        ...
        "ribbons": [
            ...
            {
                // Details of the ribbon configuration would be here.

                "requirements": {
                    "capabilities": {
                        "name": "MailBox",
                        "minVersion": "1.9"
                    }
                }
            }
        ]
    }
]
```

### extensions.runtimes.requirements

The "extensions.runtimes" property configures the sets of runtimes and actions that each extension point can use. For more information on its usage, see [create add-in commands](/office/dev/add-ins/develop/create-addin-commands-unified-manifest), [configure the runtime for a task pane](/office/dev/add-ins/develop/create-addin-commands-unified-manifest#configure-the-runtime-for-the-task-pane-command), and [configure the runtime for the function command](/office/dev/add-ins/develop/create-addin-commands-unified-manifest#configure-the-runtime-for-the-function-command). For more information about runtimes in Office add-ins, see [Runtimes in Office Add-ins](../testing/runtimes.md).

The "requirements" subproperty can be used to prevent the runtime from being included in versions of Office or in Office applications where it wouldn't be used.

The example earlier in [extensions.autoRunEvents.requirements](#extensionsautoruneventsrequirements) shows how to block the autolaunch feature in versions that don't support all of the code in the `logOutgoingEmail` function, which includes code that requires **Mailbox 1.13**. Suppose that in that same scenario, the "runtime" object that is configured to support the "logOutgoingEmail" action isn't configured to support any other action. In that case, the developer should block the runtime object in versions that don't support **Mailbox 1.13** since it would never be used. The following is an example. For details of the runtime configuration, see [Create add-in commands with the unified manifest for Microsoft 365](create-addin-commands-unified-manifest.md).


```json
"extensions": [
    ...
    {
        ...
        "runtimes": [
            ...
            {
                // Details of the runtime configuration would be here.

                "requirements": {
                    "capabilities": {
                        "name": "MailBox",
                        "minVersion": "1.13"
                    }
                }
            }
        ]
    }
]
```

Similarly, for the example in [extensions.ribbons.requirements](#extensionsribbonsrequirements), if the action linked to the custom button is the only action configured in a runtime object, then that runtime object should be blocked in the same circumstances in which the ribbon object is blocked.






