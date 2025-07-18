---
title: Manage the sensitivity label of your message or appointment in compose mode
description: Learn how to manage the sensitivity label of your message or appointment in compose mode.
ms.date: 07/17/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Manage the sensitivity label of your message or appointment in compose mode

Collaboration in the workplace not only occurs within the organization, but extends to external partners as well. With information being shared beyond an organization's network, it's important to establish measures to prevent data loss and enforce compliance policies. [Microsoft Purview Information Protection](/microsoft-365/compliance/information-protection) helps you implement solutions to classify and protect sensitive information. The use of sensitivity labels in Outlook is a capability you can configure to protect your data.

You can use the Office JavaScript API to implement sensitivity label solutions in your Outlook add-in projects and support the following scenarios.

- Automatically apply sensitivity labels to certain messages and appointments while they're being composed, so that users can focus on their work.
- Restrict additional actions if a certain sensitivity label is applied to a message or appointment, such as preventing users from adding external recipients to a message.
- Add a header or footer to a message or appointment based on its sensitivity label to comply with business and legal policies.

> [!NOTE]
> Support for the sensitivity label feature was introduced in [requirement set 1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13). For information about client support for this feature, see [Supported clients and platforms](#supported-clients-and-platforms).

## Prerequisites

To implement the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).

### Supported clients and platforms

The following table lists client-server combinations that support the use of the sensitivity label feature in Outlook add-ins. Excluded combinations aren't supported.

|Client|Exchange Online|
|------|------|
|**Web browser (modern UI)**<br><br>[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)|Supported|
|**Windows (classic)**<br>Version 2304 (Build 16327.20248) or later|Supported|
|**Mac**<br>Version 16.77 (23081600) or later|Supported|
|**Android**|Not applicable|
|**iOS**|Not applicable|

## Configure the manifest

To use the sensitivity feature in your Outlook add-in project, you must configure the **read/write item** permission in the manifest of your add-in.

- **Unified manifest for Microsoft 365**: In the [`"authorization.permissions.resourceSpecific"`](/microsoft-365/extensibility/schema/root-authorization-permissions#resourcespecific) array, set the `"name"` property of an object to `"MailboxItem.ReadWrite.User"`.
- **Add-in only manifest**: Set the [\<Permissions\> element](/javascript/api/manifest/permissions) to **ReadWriteItem**.

If your add-in will detect and handle the `OnSensitivityLabelChanged` event, additional manifest configurations are required to enable the event-based activation feature. To learn more, see [Detect sensitivity label changes with the OnSensitivityLabelChanged event](#detect-sensitivity-label-changes-with-the-onsensitivitylabelchanged-event).

[!INCLUDE [outlook-sensitivity-label-event-support](../includes/outlook-sensitivity-label-event-support.md)]

## Verify the status of the catalog of sensitivity labels

Sensitivity labels and policies are configured by an organization's administrator through the [Microsoft Purview compliance portal](/microsoft-365/compliance/microsoft-365-compliance-center). For guidance on how to configure sensitivity labels in your tenant, see [Create and configure sensitivity labels and their policies](/microsoft-365/compliance/create-sensitivity-labels).

Before you can get or set the sensitivity label on a message or appointment, you must first ensure that the catalog of sensitivity labels is enabled on the mailbox where the add-in is installed. To check the status of the catalog of sensitivity labels, call [context.sensitivityLabelsCatalog.getIsEnabledAsync](/javascript/api/outlook/office.sensitivitylabelscatalog#outlook-office-sensitivitylabelscatalog-getisenabledasync-member(1)) in compose mode.

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

## Identify available sensitivity labels

If you want to determine the sensitivity labels available for use on a message or appointment in compose mode, use [context.sensitivityLabelsCatalog.getAsync](/javascript/api/outlook/office.sensitivitylabelscatalog#outlook-office-sensitivitylabelscatalog-getasync-member(1)). The available labels are returned in the form of [SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails) objects, which provide the following details.

- The name of the label.
- The unique identifier (GUID) of the label.
- A description of the label.
- The color assigned to the label.
- The configured [sublabels](/microsoft-365/compliance/sensitivity-labels#sublabels-grouping-labels), if any.

The following example shows how to identify the sensitivity labels available in the catalog.

```javascript
// It's recommended to check the status of the catalog of sensitivity labels before
// calling other sensitivity label methods.
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

## Get the sensitivity label of a message or appointment

To get the sensitivity label currently applied to a message or appointment in compose mode, call [item.sensitivityLabel.getAsync](/javascript/api/outlook/office.sensitivitylabel#outlook-office-sensitivitylabel-getasync-member(1)) as shown in the following example. This returns the GUID of the sensitivity label.

```javascript
// It's recommended to check the status of the catalog of sensitivity labels before
// calling other sensitivity label methods.
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

## Set the sensitivity label on a message or appointment

You can set only one sensitivity label on a message or appointment in compose mode. Before you set the label, call [context.sensitivityLabelsCatalog.getAsync](/javascript/api/outlook/office.sensitivitylabelscatalog#outlook-office-sensitivitylabelscatalog-getasync-member(1)). This ensures that the label you want to apply is available for use. It also helps you identify a label's GUID, which you'll need to apply the label to the mail item. After you confirm the label's availability, pass its GUID as a parameter to [item.sensitivityLabel.setAsync](/javascript/api/outlook/office.sensitivitylabel#outlook-office-sensitivitylabel-setasync-member(1)), as shown in the following example.

```javascript
// It's recommended to check the status of the catalog of sensitivity labels before
// calling other sensitivity label methods.
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

Instead of using the GUID to set the sensitivity label, you can pass the [SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails) object retrieved from the catalog call, as shown in the following example.

```javascript
// It's recommended to check the status of the catalog of sensitivity labels before
// calling other sensitivity label methods.
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

## Detect sensitivity label changes with the OnSensitivityLabelChanged event

Take extra measures to protect your data by using the `OnSensitivityLabelChanged` event. This event enables your add-in to complete tasks in response to sensitivity label changes on a message or appointment. For example, you can prevent users from downgrading the sensitivity label of a mail item if it contains certain attachments.

The `OnSensitivityLabelChanged` event is available through the event-based activation feature. To learn how to configure, debug, and deploy an event-based add-in that uses this event, see [Activate add-ins with events](../develop/event-based-activation.md).

[!INCLUDE [outlook-sensitivity-label-event-support](../includes/outlook-sensitivity-label-event-support.md)]

## See also

- [Learn about sensitivity labels](/microsoft-365/compliance/sensitivity-labels)
- [Get started with sensitivity labels](/microsoft-365/compliance/get-started-with-sensitivity-labels)
- [Create and configure sensitivity labels and their policies](/microsoft-365/compliance/create-sensitivity-labels)
- [Activate add-ins with events](../develop/event-based-activation.md)
- [Office Add-ins code sample: Verify the sensitivity label of a message](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-verify-sensitivity-label)
