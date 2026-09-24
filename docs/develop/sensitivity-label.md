---
title: Manage sensitivity labels in Office Add-ins
description: Learn how to manage sensitivity labels in Excel, Outlook, PowerPoint, and Word add-ins.
ms.date: 09/23/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Manage sensitivity labels in Office Add-ins

Workplace collaboration often extends beyond an organization to external partners. Sharing information outside an organization's network requires measures to prevent data loss and enforce compliance policies. [Microsoft Purview Information Protection](/microsoft-365/compliance/information-protection) provides solutions for classifying and protecting sensitive information. Sensitivity labels apply this protection to data in Excel, Outlook, PowerPoint, and Word.

Use the Office JavaScript API to implement sensitivity label solutions in your Office Add-in projects and support the following scenarios.

- Apply sensitivity labels to documents, messages, or appointments to comply with business and legal policies.
- Restrict additional actions if a certain sensitivity label is applied, such as preventing users from adding external recipients to a message.
- Classify data based on its sensitivity label to support auditing and reporting.

> [!NOTE]
> In Excel, PowerPoint, and Word, the sensitivity label APIs are in preview. In Outlook, support for the sensitivity label feature was introduced in [requirement set 1.13](/javascript/api/requirement-sets/outlook/outlook-requirement-set-1-13). For information about client support, see [Supported clients and platforms](#supported-clients-and-platforms).

## Prerequisites

The sensitivity label feature requires a Microsoft 365 E5 subscription. Check whether you qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) in the [program FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Otherwise, [start a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).

### Supported clients and platforms

Sensitivity label API support varies by Office application and platform. For details, see the following sections.

- [Excel, PowerPoint, and Word](#excel-powerpoint-and-word)
- [Outlook](#outlook)

#### Excel, PowerPoint, and Word

The sensitivity label APIs are in preview in Excel, PowerPoint, and Word on Windows, Mac, and the web.

#### Outlook

The following table lists client-server combinations that support the use of the sensitivity label feature in Outlook add-ins. Excluded combinations aren't supported.

|Client|Exchange Online|
|------|------|
|**Web browser (modern UI)**<br><br>[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)|Supported|
|**Windows (classic)**<br>Version 2304 (Build 16327.20248) or later|Supported|
|**Mac**<br>Version 16.77 (23081600) or later|Supported|
|**Android**|Not applicable|
|**iOS**|Not applicable|

## Configure sensitivity label support

# [Excel, PowerPoint, and Word](#tab/excel-powerpoint-word)

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

The Excel, PowerPoint, and Word sensitivity label APIs follow a similar programming pattern. In each host, the request context provides access to the sensitivity label catalog, while the host-specific file object provides methods to get or update its label.

The following table lists the API members used to access the sensitivity label catalog and the label applied to a file in each Office host application.

|Application|Sensitivity label catalog|Sensitivity label on the file|
|---|---|---|
|Excel|[`context.sensitivityLabelsCatalog`](/javascript/api/excel/excel.requestcontext?view=excel-js-preview&preserve-view=true#excel-excel-requestcontext-sensitivitylabelscatalog-member)|[`context.workbook.sensitivityLabel`](/javascript/api/excel/excel.workbook?view=excel-js-preview&preserve-view=true#excel-excel-workbook-sensitivitylabel-member)|
|PowerPoint|[`context.sensitivityLabelsCatalog`](/javascript/api/powerpoint/powerpoint.requestcontext?view=powerpoint-js-preview&preserve-view=true#powerpoint-powerpoint-requestcontext-sensitivitylabelscatalog-member)|[`context.presentation.sensitivityLabel`](/javascript/api/powerpoint/powerpoint.presentation?view=powerpoint-js-preview&preserve-view=true#powerpoint-powerpoint-presentation-sensitivitylabel-member)|
|Word|[`context.sensitivityLabelsCatalog`](/javascript/api/word/word.requestcontext?view=word-js-preview&preserve-view=true#word-word-requestcontext-sensitivitylabelscatalog-member)|[`context.document.sensitivityLabel`](/javascript/api/word/word.document?view=word-js-preview&preserve-view=true#word-word-document-sensitivitylabel-member)|

The examples in the following sections use Word. To use Excel or PowerPoint, substitute the corresponding host namespace and file-level sensitivity label object.

# [Outlook](#tab/outlook)

To use the sensitivity label feature in an Outlook add-in, you must configure the **read/write item** permission in the manifest of your add-in.

- **Unified manifest for Microsoft 365**: In the [`"authorization.permissions.resourceSpecific"`](/microsoft-365/extensibility/schema/root-authorization-permissions#resourcespecific) array, set the `"name"` property of an object to `"MailboxItem.ReadWrite.User"`.
- **Add-in only manifest**: Set the [\<Permissions\> element](/javascript/api/manifest/permissions) to **ReadWriteItem**.

[!INCLUDE [outlook-unified-manifest-mac](../includes/outlook-unified-manifest-mac.md)]

Add-ins that handle the `OnSensitivityLabelChanged` event require additional manifest configuration for event-based activation. For details, see [Detect sensitivity label changes with the OnSensitivityLabelChanged event](#detect-sensitivity-label-changes-with-the-onsensitivitylabelchanged-event).

---

## Verify sensitivity labeling is available

Sensitivity labels and policies are configured by an organization's administrator through the [Microsoft Purview compliance portal](/microsoft-365/compliance/microsoft-365-compliance-center). For guidance on how to configure sensitivity labels in your tenant, see [Create and configure sensitivity labels and their policies](/microsoft-365/compliance/create-sensitivity-labels).

# [Excel, PowerPoint, and Word](#tab/excel-powerpoint-word)

To determine whether sensitivity labeling is available to the current user, load `getLabelingCapability` ([Excel](/javascript/api/excel/excel.sensitivitylabelscatalog?view=excel-js-preview&preserve-view=true#excel-excel-sensitivitylabelscatalog-getlabelingcapability-member), [PowerPoint](/javascript/api/powerpoint/powerpoint.sensitivitylabelscatalog?view=powerpoint-js-preview&preserve-view=true#powerpoint-powerpoint-sensitivitylabelscatalog-getlabelingcapability-member), [Word](/javascript/api/word/word.sensitivitylabelscatalog?view=word-js-preview&preserve-view=true#word-word-sensitivitylabelscatalog-getlabelingcapability-member)) from the sensitivity label catalog.

```typescript
await Word.run(async (context) => {
    // Access the sensitivity label catalog for the current user.
    const labelCatalog = context.sensitivityLabelsCatalog;
    if (!labelCatalog) {
        console.warn("The sensitivity label catalog isn't available.");
        return;
    }

    // Load the labeling capability status before reading it.
    labelCatalog.load("getLabelingCapability");
    await context.sync();

    // Display whether sensitivity labeling is enabled and available.
    console.log(`Sensitivity labeling capability: ${labelCatalog.getLabelingCapability}`);
});
```

# [Outlook](#tab/outlook)

Before getting or setting the sensitivity label on a message or appointment, verify that the sensitivity label catalog is enabled on the mailbox where the add-in is installed. Call [context.sensitivityLabelsCatalog.getIsEnabledAsync](/javascript/api/outlook/office.sensitivitylabelscatalog#outlook-office-sensitivitylabelscatalog-getisenabledasync-member(1)) in compose mode.

```javascript
// Check whether the catalog of sensitivity labels is enabled.
Office.context.sensitivityLabelsCatalog.getIsEnabledAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log(asyncResult.value);
    } else {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
});
```

---

## Identify available sensitivity labels

# [Excel, PowerPoint, and Word](#tab/excel-powerpoint-word)

To retrieve the labels published to the current user, call `getLabels()` ([Excel](/javascript/api/excel/excel.sensitivitylabelscatalog?view=excel-js-preview&preserve-view=true#excel-excel-sensitivitylabelscatalog-getlabels-member(1)), [PowerPoint](/javascript/api/powerpoint/powerpoint.sensitivitylabelscatalog?view=powerpoint-js-preview&preserve-view=true#powerpoint-powerpoint-sensitivitylabelscatalog-getlabels-member(1)), [Word](/javascript/api/word/word.sensitivitylabelscatalog?view=word-js-preview&preserve-view=true#word-word-sensitivitylabelscatalog-getlabels-member(1))) on the catalog.

The method returns a collection whose items and properties aren't available until you explicitly load them and call `context.sync()`. Load `items` ([Excel](/javascript/api/excel/excel.sensitivitylabeldetailscollection?view=excel-js-preview&preserve-view=true#excel-excel-sensitivitylabeldetailscollection-items-member), [PowerPoint](/javascript/api/powerpoint/powerpoint.sensitivitylabeldetailscollection?view=powerpoint-js-preview&preserve-view=true#powerpoint-powerpoint-sensitivitylabeldetailscollection-items-member), [Word](/javascript/api/word/word.sensitivitylabeldetailscollection?view=word-js-preview&preserve-view=true#word-word-sensitivitylabeldetailscollection-items-member)) and the label properties your add-in needs. Available properties differ by host. For a complete list, see `SensitivityLabelDetails` ([Excel](/javascript/api/excel/excel.sensitivitylabeldetails?view=excel-js-preview&preserve-view=true), [PowerPoint](/javascript/api/powerpoint/powerpoint.sensitivitylabeldetails?view=powerpoint-js-preview&preserve-view=true), [Word](/javascript/api/word/word.sensitivitylabeldetails?view=word-js-preview&preserve-view=true)).

```typescript
await Word.run(async (context) => {
    // Access the sensitivity label catalog for the current user.
    const labelCatalog = context.sensitivityLabelsCatalog;
    if (!labelCatalog) {
        console.warn("The sensitivity label catalog isn't available.");
        return;
    }

    // Get the available labels and load the properties used by the add-in.
    const availableLabels = labelCatalog.getLabels();
    availableLabels.load("items/id,items/name,items/isEnabled");
    await context.sync();

    // Display the available labels.
    console.log("Available sensitivity labels:");
    availableLabels.items.forEach((label) => {
        console.log(`${label.name} (${label.id}) - ${label.isEnabled ? "Enabled" : "Disabled"}`);
    });
});
```

# [Outlook](#tab/outlook)

To determine the sensitivity labels available for use on a message or appointment in compose mode, use [context.sensitivityLabelsCatalog.getAsync](/javascript/api/outlook/office.sensitivitylabelscatalog#outlook-office-sensitivitylabelscatalog-getasync-member(1)). The available labels are returned in the form of [SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails) objects, which provide the following details.

- The name of the label.
- The unique identifier (GUID) of the label.
- A description of the label.
- The color assigned to the label.
- The configured [sublabels](/microsoft-365/compliance/sensitivity-labels#sublabels-grouping-labels), if any.

The following example shows how to identify the sensitivity labels available in the catalog.

```javascript
// Check the sensitivity label catalog status before calling other label methods.
Office.context.sensitivityLabelsCatalog.getIsEnabledAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value == true) {
        // Identify available sensitivity labels in the catalog.
        Office.context.sensitivityLabelsCatalog.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const catalog = asyncResult.value;
                console.log("Sensitivity Labels Catalog:");
                catalog.forEach((sensitivityLabel) => {
                    console.log(`Name: ${sensitivityLabel.name}`);
                    console.log(`ID: ${sensitivityLabel.id}`);
                    console.log(`Tooltip: ${sensitivityLabel.tooltip}`);
                    console.log(`Color: ${sensitivityLabel.color}`);
                    console.log(`Sublabels: ${JSON.stringify(sensitivityLabel.children)}`);
                });
            } else {
                console.log("Action failed with error: " + asyncResult.error.message);
            }
        });
    } else {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
});
```

---

## Get the sensitivity label

# [Excel, PowerPoint, and Word](#tab/excel-powerpoint-word)

To retrieve the current label, if one is applied, call `getCurrentOrNullObject()` ([Excel](/javascript/api/excel/excel.sensitivitylabel?view=excel-js-preview&preserve-view=true#excel-excel-sensitivitylabel-getcurrentornullobject-member(1)), [PowerPoint](/javascript/api/powerpoint/powerpoint.sensitivitylabel?view=powerpoint-js-preview&preserve-view=true#powerpoint-powerpoint-sensitivitylabel-getcurrentornullobject-member(1)), [Word](/javascript/api/word/word.sensitivitylabel?view=word-js-preview&preserve-view=true#word-word-sensitivitylabel-getcurrentornullobject-member(1))) on the file's sensitivity label object.

```typescript
await Word.run(async (context) => {
    // Access the sensitivity label applied to the current document.
    const documentLabel = context.document.sensitivityLabel;

    // Get the current label, if one is applied, and load its ID and name.
    const currentLabel = documentLabel.getCurrentOrNullObject();
    currentLabel.load("id,name");
    await context.sync();

    // Display the current label or report that the document isn't labeled.
    if (currentLabel.isNullObject) {
        console.log("The document doesn't have a sensitivity label.");
    } else {
        console.log(`Current label: ${currentLabel.name} (${currentLabel.id})`);
    }
});
```

# [Outlook](#tab/outlook)

To get the sensitivity label currently applied to a message or appointment in compose mode, call [item.sensitivityLabel.getAsync](/javascript/api/outlook/office.sensitivitylabel#outlook-office-sensitivitylabel-getasync-member(1)) as shown in the following example. This returns the GUID of the sensitivity label.

```javascript
// Check the sensitivity label catalog status before calling other label methods.
Office.context.sensitivityLabelsCatalog.getIsEnabledAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value == true) {
        // Get the current sensitivity label of a message or appointment.
        Office.context.mailbox.item.sensitivityLabel.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log(asyncResult.value);
            } else {
                console.log("Action failed with error: " + asyncResult.error.message);
            }
        });
    } else {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
});
```

---

## Set the sensitivity label

# [Excel, PowerPoint, and Word](#tab/excel-powerpoint-word)

Before applying a label, call `getLabels()` ([Excel](/javascript/api/excel/excel.sensitivitylabelscatalog?view=excel-js-preview&preserve-view=true#excel-excel-sensitivitylabelscatalog-getlabels-member(1)), [PowerPoint](/javascript/api/powerpoint/powerpoint.sensitivitylabelscatalog?view=powerpoint-js-preview&preserve-view=true#powerpoint-powerpoint-sensitivitylabelscatalog-getlabels-member(1)), [Word](/javascript/api/word/word.sensitivitylabelscatalog?view=word-js-preview&preserve-view=true#word-word-sensitivitylabelscatalog-getlabels-member(1))) and select an enabled label or sublabel from the returned collection. The `tryToUpdate()` method ([Excel](/javascript/api/excel/excel.sensitivitylabel?view=excel-js-preview&preserve-view=true#excel-excel-sensitivitylabel-trytoupdate-member(1)), [PowerPoint](/javascript/api/powerpoint/powerpoint.sensitivitylabel?view=powerpoint-js-preview&preserve-view=true#powerpoint-powerpoint-sensitivitylabel-trytoupdate-member(1)), [Word](/javascript/api/word/word.sensitivitylabel?view=word-js-preview&preserve-view=true#word-word-sensitivitylabel-trytoupdate-member(1))) requires the selected label's ID as its parameter. Calling `getLabels()` first lets you retrieve this required ID and verify that the label is available to the current user. Check the returned `SensitivityLabelUpdateResult` ([Excel](/javascript/api/excel/excel.sensitivitylabelupdateresult?view=excel-js-preview&preserve-view=true), [PowerPoint](/javascript/api/powerpoint/powerpoint.sensitivitylabelupdateresult?view=powerpoint-js-preview&preserve-view=true), [Word](/javascript/api/word/word.sensitivitylabelupdateresult?view=word-js-preview&preserve-view=true)) value to determine whether the update succeeded.

> [!NOTE]
> A parent label that has sublabels can't be applied directly. Select one of its enabled sublabels instead.

```typescript
async function setDocumentSensitivityLabel(labelId: string) {
    await Word.run(async (context) => {
        // Access the sensitivity label applied to the current document.
        const documentLabel = context.document.sensitivityLabel;

        // Apply the selected label.
        const updateResult = documentLabel.tryToUpdate(labelId);
        await context.sync();

        // Check whether the label update succeeded.
        if (updateResult.value === Word.SensitivityLabelUpdateResult.success) {
            console.log("Applied the sensitivity label to the document.");
        } else {
            console.error(`The sensitivity label wasn't applied. Result: ${updateResult.value}`);
        }
    });
}
```

# [Outlook](#tab/outlook)

Outlook supports one sensitivity label on a message or appointment in compose mode. Before setting the label, call [context.sensitivityLabelsCatalog.getAsync](/javascript/api/outlook/office.sensitivitylabelscatalog#outlook-office-sensitivitylabelscatalog-getasync-member(1)) to verify that the label is available and retrieve its GUID. Then, pass the GUID to [item.sensitivityLabel.setAsync](/javascript/api/outlook/office.sensitivitylabel#outlook-office-sensitivitylabel-setasync-member(1)), as shown in the following example.

```javascript
// Check the sensitivity label catalog status before calling other label methods.
Office.context.sensitivityLabelsCatalog.getIsEnabledAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value == true) {
        // Identify available sensitivity labels in the catalog.
        Office.context.sensitivityLabelsCatalog.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const catalog = asyncResult.value;
                if (catalog.length > 0) {
                    // Get the GUID of the sensitivity label.
                    var id = catalog[0].id;
                    // Set the mail item's sensitivity label using the label's GUID.
                    Office.context.mailbox.item.sensitivityLabel.setAsync(id, (asyncResult) => {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log(asyncResult.status);
                        } else {
                            console.log("Action failed with error: " + asyncResult.error.message);
                        }
                    });
                } else {
                    console.log("Catalog list is empty");
                }
            } else {
                console.log("Action failed with error: " + asyncResult.error.message);
            }
        });
    } else {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
});
```

Alternatively, pass the [SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails) object returned by the catalog call, as shown in the following example.

```javascript
// Check the sensitivity label catalog status before calling other label methods.
Office.context.sensitivityLabelsCatalog.getIsEnabledAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value == true) {
        // Identify available sensitivity labels in the catalog.
        Office.context.sensitivityLabelsCatalog.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const catalog = asyncResult.value;
                if (catalog.length > 0) {
                    // Set the mail item's sensitivity label using the SensitivityLabelDetails object.
                    Office.context.mailbox.item.sensitivityLabel.setAsync(catalog[0], (asyncResult) => {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log(asyncResult.status);
                        } else {
                            console.log("Action failed with error: " + asyncResult.error.message);
                        }
                    });
                } else {
                    console.log("Catalog list is empty");
                }
            } else {
                console.log("Action failed with error: " + asyncResult.error.message);
            }
        });
    } else {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
});
```

---

## Detect sensitivity label changes with the OnSensitivityLabelChanged event

> [!NOTE]
> The `OnSensitivityLabelChanged` event is only available in Outlook.

Use the `OnSensitivityLabelChanged` event to run add-in logic when the sensitivity label changes on a message or appointment. For example, prevent users from downgrading the label of a mail item that contains certain attachments.

The `OnSensitivityLabelChanged` event uses event-based activation. For configuration, debugging, and deployment guidance, see [Activate add-ins with events](event-based-activation.md).

## See also

- [Learn about sensitivity labels](/microsoft-365/compliance/sensitivity-labels)
- [Get started with sensitivity labels](/microsoft-365/compliance/get-started-with-sensitivity-labels)
- [Create and configure sensitivity labels and their policies](/microsoft-365/compliance/create-sensitivity-labels)
- [Activate add-ins with events](event-based-activation.md)
- [Office Add-ins code sample: Verify the sensitivity label of a message](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-verify-sensitivity-label)
