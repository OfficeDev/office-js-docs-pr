---
title: Manage both a unified manifest and an XML manifest version of your Office Add-in
description: Learn when and how to maintain versions of your add-in for each type of manifest.
ms.topic: best-practice
ms.date: 07/13/2023
ms.localizationpriority: medium
---

# Manage both a unified manifest and an XML manifest version of your Office Add-in

Microsoft is making a number of improvements to the Microsoft 365 developer platform. These improvements provide more consistency in the development, deployment, installation, and administration of all types of extensions of Microsoft 365, including Office Add-ins.

One important improvement we're working on is the ability to create a single unit of distribution for all your Microsoft 365 extensions by using the same manifest format: the JSON-formatted [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md).

> [!TIP]
> For information about converting an existing add-in to use the unified manifest, see [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md).

After you've created a version of your add-in that uses the unified manifest, you must ensure the following:

- The new version is the one that users see in all Office versions that support the unified manifest.
- The old version is still available in versions that don't support the unified manifest.
- Users never see two seemingly identical versions; that is, the same name and icons twice in the Office UI.

To accomplish these goals, you need to configure the unified manifest to identify the existing add-in. (Don't remove the existing add-in from AppSource or the Microsoft 365 admin center.) After the new add-in version is deployed, what users will see depends on whether the Office version they are working in supports the unified manifest.

- If it doesn't support the unified manifest, they will see and work with the existing version of the add-in.
- If it does support the unified manifest, when a user runs the add-in, they'll see a prompt to install the new version. If they choose not to, the existing add-in runs. If they choose to upgrade, the existing add-in will be hidden.   

To configure the unified manifest: 

1. Scroll to the extension object in the "extensions" array.  
1. Create an "alternatives" array property, if there isn’t one already. 
1. In the "alternatives" array, create an object that has a "hide" property. 
1. Give the "hide" object a "storeOfficeAddin" property. 
1. Give the "storeOfficeAddin" object an "officeAddinId" property with the GUID of the existing add-in as its value. The following is an example.

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

1. If you're marketing the add-in through [AppSource](https://appsource.microsoft.com/), there's a further step. Give the "customOfficeAddin" property an additional child property named "assetId" with the AppSource asset ID as its value. The following is an example.

    ```json
    "hide": {
        "customOfficeAddin": {
            "officeAddinId": "b5a2794d-4aa5-4023-a84b-c60a3cbd33d4",
            "assetId": "WA999999999"
        }
    }
    ```

## Maintain both versions for the immediate future

When all of your users are working with Office versions that support the unified manifest, you can remove the XML version from deployment, but you need to maintain both versions for the immediate future.

Initially add-ins that support the unified manifest can only be installed on Microsoft 365 Version 2309 (Build ?????.?????) and later. We're working hard to bring support to older versions of Microsoft 365. In the meantime, you need to maintain both versions.

There are also some scenarios where you might want to maintain both both versions of the add-in for an extended period. For example, there are two features of add-ins that aren't supported with the unified manifest because they are little used or deprecated. You may choose to maintain a version of your add-in that uses these features. The following are the features that aren't supported in the unified manifest.

- [Outlook modules](../outlook/extension-module-outlook-add-ins.md) aren't supported. (But you can provide a nearly identical experience using the unified manifest by [combining an Outlook add-in and a Teams Tab in a single app](https://github.com/OfficeDev/TeamsFx/wiki/Configure-Outlook-Add-in-capability-within-your-Teams-app).)
- [Outlook contextual add-ins](../outlook/contextual-outlook-add-ins.md) (also known as "activation rules") aren't supported. (But you can provide similar experiences using the unified manifest and [Event-based activation](../outlook/autolaunch.md).)

The critical requirement for making two versions available is to be sure that the two of them appear distinct in the Outlook UI. 

- Give the new version a different name from the existing add-in. 
- Create and use different icons for the new version.
- Be sure that the "id" property of the unified manifest in the new version is a different GUID from the **\<Id\>** element in the XML manifest of the existing add-in.

> [!NOTE]
> If you use the same name and icon, the old and new solutions appear indistinguishable in the Outlook UI for add-in installation. 

