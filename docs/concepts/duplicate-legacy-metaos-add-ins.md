---
title: Managing both a unified manifest and an XML manifest version of your Office Add-in
description: Learn when and how to maintain versions of your add-in for each type of manifest.
ms.topic: best-practice
ms.date: 07/19/2023
ms.localizationpriority: medium
---

# Managing both a unified manifest and an XML manifest version of your Office Add-in

After you've created a version of your add-in that uses the [unified manifest for Microsoft 365](../develop/json-manifest-overview.md), you need to decide whether to maintain the existing add-in and, if so, what relationship it will have with the new version. In the long run, the best practice is to replace the existing add-in with the new version, but initially add-ins that support the unified manifest can only be installed on Microsoft 365 version 2309 ?????.????? and later. We're working hard to bring support to older versions of Microsoft 365. In the meantime, you need to maintain both versions.

> [!TIP]
> For information about converting an existing add-in to use the unified manifest, see [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md).

There are also some scenarios where you need to maintain both for a longer period, including the following.

- Even after support for the unified manifest is available in older versions of Microsoft 365, there will be a further period in which add-ins that use it won't be installable on perpetual versions of Office, such as volume-licensed perpetual versions of Office 2021 (and earlier).
- There are some features of add-ins that are little used or deprecated. These aren't supported with the unified add-in. Your may choose to maintain a version of your add-in that uses these features. The following are examples.

    - [Outlook modules](../outlook/extension-module-outlook-add-ins.md) aren't supported. (But you can provide a nearly identical experience using the unified manifest by [combining an Outlook add-in and a Teams Tab in a single app](https://github.com/OfficeDev/TeamsFx/wiki/Configure-Outlook-Add-in-capability-within-your-Teams-app).)
    - [Outlook contextual add-ins](../outlook/contextual-outlook-add-ins.md) (aka "activation rules") aren't supported. (But you can provide similar experiences using the unified manifest and [Event-based activation](../outlook/autolaunch.md).)

## Maintain both versions

The critical requirement for maintaining two versions is to be sure that the two of them appear distinct in the Outlook UI. 

- Give the new version a different name from the existing add-in. 
- Create and use different icons for the new version.
- Be sure that the "id" property of the unified manifest in the new version is a different GUID from the **\<Id\>** element in the XML manifest of the existing add-in.

> [!NOTE]
> If you use the same name and icon, the old and new solutions appear indistinguishable in the Outlook UI for add-in installation. 

## Replace the existing add-in

When you're ready to replace existing add-in, you need to configure the unified manifest to identify the existing add-in. (Don't remove the existing add-in from AppSource or the Microsoft 365 admin center.) After the new add-in is deployed, when a user runs the existing add-in, they'll see a prompt to install the new version. If they choose not to, the existing add-in runs. If they choose to upgrade, the existing add-in will be hidden.   

To configure the unified manifest: 

1. Open the extension object in the “extensions” array.  
1. Create an “alternatives” array property, if there isn’t one already. 
1. In the “alternatives” array, create an object that has a “hide” property. 
1. Give the “hide” object a “storeOfficeAddin” property. 
1. Give the “storeOfficeAddin” object an “officeAddinId” property with the GUID of the existing add-in as its value. The following is an example:

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
                        "customOfficeAddin": {
                            "officeAddinId": "b5a2794d-4aa5-4023-a84b-c60a3cbd33d4"
                        }
                    }
                }
            ]
        }
    ]
    ```

1. If you are marketing the add-in through [AppSource](https://appsource.microsoft.com/) there is a further step. Give the "customOfficeAddin" property an additional child property named "assetId" with the AppSource asset ID as its value. The following is an example: 

    ```json
    "hide": {
        "customOfficeAddin": {
            "officeAddinId": "b5a2794d-4aa5-4023-a84b-c60a3cbd33d4",
            "assetId": "WA999999999"
        }
    }
    ```
