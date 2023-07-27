---
title: Manage both a unified manifest and an XML manifest version of your Office Add-in
description: Learn when and how to maintain versions of your add-in for each type of manifest.
ms.topic: best-practice
ms.date: 08/01/2023
ms.localizationpriority: medium
---

# Manage both a unified manifest and an XML manifest version of your Office Add-in

Microsoft is making a number of improvements to the Microsoft 365 developer platform. These improvements provide more consistency in the development, deployment, installation, and administration of all types of extensions of Microsoft 365, including Office Add-ins.

One important improvement we're working on is the ability to create a single unit of distribution for all your Microsoft 365 extensions by using the same manifest format: the JSON-formatted [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md).

> [!TIP]
> For information about converting an existing add-in to use the unified manifest, see [Convert an add-in to use the unified manifest for Microsoft 365](../develop/convert-xml-to-json-manifest.md).

After you've created a version of your add-in that uses the unified manifest, you must link the existing add-in and the app built using the unified manifest to ensure users never see two copies of the add-in UI inside of Outlook. Use the following steps.

1. Open the extension object in the “extensions” array.
1. Create an “alternatives” array property, if there isn’t one already.
1. In the “alternatives” array, create an “alternate” object that has a “hide” property.
1. If the existing add-in is marketed through AppSource, give the “hide” object a “storeOfficeAddin” property. Otherwise, skip to step 6.
1. Give the “storeOfficeAddin” object two properties:

    - An “officeAddinId” with the GUID of the old add-in as its value.
    - An “assetId” with the AppSource asset ID as its value.

    The following is an example:

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
                    }
                }
            ]
        }
    ]
    ```

    > [!NOTE]
    > 
    > - The asset ID of the add-in in your unified manifest must match with an existing add-in that has been published by your seller account on Partner Center. If the asset ID of the add-in that you have linked in your unified manifest does not match an existing offer published by your seller account, the unified manifest submission will fail.  You will need to update the right add-in asset ID and re-submit the unified manifest. 
    > - An existing add-in can only be hidden by a single unified manifest. At this time, you cannot use multiple unified manifests to hide the same add-in. If you try to hide an already linked add-in using a different unified manifest, the submission will fail. You will need to remove the linking and re-submit the unified manifest.


1. If the old add-in isn't distributed through AppSource, then give the “hide” object a “customOfficeAddin” property.
1. Give the “customOfficeAddin” object an “officeAddinId” property with the GUID of the old add-in as its value. The following is an example:

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

Don't remove the existing add-in from AppSource or the Microsoft 365 Admin Center, or earlier versions of Office will no longer be able to use your add-in.

## Maintain both versions for the immediate future

When all of your users are working with Office versions that support the unified manifest, you can remove the XML version from deployment, but you need to maintain both versions for the immediate future.

Initially add-ins that support the unified manifest can only be installed on Microsoft 365 Version 2307 (Build 16626.20132) and later. We're working hard to bring support to older versions of Microsoft 365. In the meantime, you need to maintain both versions.

There are also some scenarios where you might want to maintain both both versions of the add-in for an extended period. For example, there are two features of add-ins that aren't supported with the unified manifest because they are little used or deprecated. You may choose to maintain a version of your add-in that uses these features. The following are the features that aren't supported in the unified manifest.

- [Outlook modules](../outlook/extension-module-outlook-add-ins.md) aren't supported. (But you can provide a nearly identical experience using the unified manifest by [combining an Outlook add-in and a Teams Tab in a single app](https://github.com/OfficeDev/TeamsFx/wiki/Configure-Outlook-Add-in-capability-within-your-Teams-app).)
- [Outlook contextual add-ins](../outlook/contextual-outlook-add-ins.md) (also known as "activation rules") aren't supported. (But you can provide similar experiences using the unified manifest and [Event-based activation](../outlook/autolaunch.md).)

The critical requirement for making two versions available is to be sure that the two of them appear distinct in the Outlook UI. 

- Give the new version a different name from the existing add-in. 
- Create and use different icons for the new version.
- Be sure that the "id" property of the unified manifest in the new version is a different GUID from the **\<Id\>** element in the XML manifest of the existing add-in.

> [!NOTE]
> If you use the same name and icon, the old and new solutions appear indistinguishable in the Outlook UI for add-in installation. 

