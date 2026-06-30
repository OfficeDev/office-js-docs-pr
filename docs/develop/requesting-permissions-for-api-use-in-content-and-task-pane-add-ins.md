---
title: Requesting permissions for API use in add-ins
description: Learn about different permission levels to declare in the manifest of an add-in to specify the level of JavaScript API access.
ms.date: 06/30/2026
ms.localizationpriority: medium
---

# Requesting permissions for API use in add-ins

This article describes the different permission levels that you declare in your add-in's manifest to specify the level of JavaScript API access your add-in requires for its features.

> [!IMPORTANT]
> This article applies to only non-Outlook add-ins. To learn about permission levels for Outlook add-ins, see [Understanding Outlook add-in permissions](../outlook/understanding-outlook-add-in-permissions.md).

## Permissions model

A five-level JavaScript API access-permissions model provides the basis for privacy and security for users of your add-ins. The following figure shows the five levels of API permissions you can declare in your add-in's manifest.

:::image type="content" source="../images/office15-app-sdk-task-pane-app-permission.png" alt-text="Levels of permissions for add-ins.":::

These permissions specify the subset of the API that the add-in [runtime](../testing/runtimes.md) allows your add-in to use when a user inserts, and then activates (trusts) your add-in. To declare the permission level your add-in requires, specify one of the permission values in the manifest. The markup varies depending on the type of manifest.

|Permission canonical name|Add-in only manifest name|Unified manifest name|
|:-----|:-----|:-----|
|**restricted**|Restricted|Document.Restricted.User|
|**read document**|ReadDocument|Document.Read.User|
|**read all document**|ReadAllDocument|Document.ReadAll.User|
|**write document**|WriteDocument|Document.Write.User|
|**read/write document**|ReadWriteDocument|Document.ReadWrite.User|

> [!IMPORTANT]
> If your add-in uses the [application-specific APIs](application-specific-api-model.md), declare the **read/write document** permission in the manifest. This requirement applies even when your code only reads data.

### Manifest declaration examples

- **Unified manifest for Microsoft 365**: Use the [`"authorization.permissions.resourceSpecific"`](/microsoft-365/extensibility/schema/root-authorization-permissions#resourcespecific) property. The following example requests the **read/write document** permission.

   ```json
   "authorization": {
      "permissions": {
        "resourceSpecific": [
          ...
          {
            "name": "Document.ReadWrite.User",
            "type": "Delegated"
          },
        ]
      }  
   },
   ```

- **Add-in only manifest**: Use the [Permissions](/javascript/api/manifest/permissions) element of the manifest. The following example requests the **read/write document** permission.

   ```XML
   <Permissions>ReadWriteDocument</Permissions>
   ```

## Permission levels

The following table describes the subsets of the [Common JavaScript APIs](understand-the-javascript-api-for-office.md#api-models) that are enabled by each permission level. The [application-specific APIs](application-specific-api-model.md) are not controlled by this table; they always require the **read/write document** permission, even when your add-in only reads data.

|Permission canonical name|Add-in only manifest name|Unified manifest name|Enabled subset of the Common APIs|
|:-----|:-----|:-----|:-----|
|**restricted**|Restricted|Document.Restricted.User|The methods of the [Settings](/javascript/api/office/office.settings) object, and the [Document.getActiveViewAsync](/javascript/api/office/office.document#office-office-document-getactiveviewasync-member(1)) method. This is the minimum permission level that can be requested by an add-in.|
|**read document**|ReadDocument|Document.Read.User|In addition to the API allowed by the **restricted** permission, adds access to the API members necessary to read the document and manage bindings.|
|**read all document**|ReadAllDocument|Document.ReadAll.User|In addition to the APIs allowed by the **restricted** and **read document** permissions, this level allows [Document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) and [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) methods.|
|**write document**|WriteDocument|Document.Write.User|In addition to the API allowed by the **restricted** permission, adds access to the following API members.<ul><li>The [Document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) method to write to the user's selection in the document.</li></ul>|
|**read/write document**|ReadWriteDocument|Document.ReadWrite.User|In addition to the API allowed by the **restricted**, **read document**, **read all document**, and **write document** permissions, includes access to all remaining API supported by add-ins, including methods for subscribing to events. You must declare the **read/write document** permission to access these additional API members:<ul><li><p>The [Binding.setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)) method for writing to bound regions of the document</li><li>The [TableBinding.addRowsAsync](/javascript/api/office/office.tablebinding#office-office-tablebinding-addrowsasync-member(1)) method for adding rows to bound tables.</li><li>The [TableBinding.addColumnsAsync](/javascript/api/office/office.tablebinding#office-office-tablebinding-addcolumnsasync-member(1)) method for adding columns to bound tables.</li><li>The [TableBinding.deleteAllDataValuesAsync](/javascript/api/office/office.tablebinding#office-office-tablebinding-deletealldatavaluesasync-member(1)) method for deleting all data in a bound table.</li><li>The [setFormatsAsync](/javascript/api/office/office.tablebinding#office-office-tablebinding-setformatsasync-member(1)), [clearFormatsAsync](/javascript/api/office/office.tablebinding#office-office-tablebinding-clearformatsasync-member(1)), and [setTableOptionsAsync](/javascript/api/office/office.tablebinding#office-office-tablebinding-settableoptionsasync-member(1)) methods of the TableBinding object for setting formatting and options on bound tables.</li><li>All of the members of the [CustomXmlNode](/javascript/api/office/office.customxmlnode), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [CustomXmlParts](/javascript/api/office/office.customxmlparts), and [CustomXmlPrefixMappings](/javascript/api/office/office.customxmlprefixmappings) objects</li><li>All of the methods for subscribing to the events supported by add-ins, specifically the `addHandlerAsync` and `removeHandlerAsync` methods of the [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [ProjectDocument](/javascript/api/office/office.document), and [Settings](/javascript/api/office/office.document#office-office-document-settings-member) objects.</li></ul>|

## See also

- [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md)
