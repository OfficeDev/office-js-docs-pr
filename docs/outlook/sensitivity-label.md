---
title: Manage the sensitivity label of your message or appointment in compose mode (preview)
description: Learn how to manage the sensitivity label of your message or appointment in compose mode.
ms.date: 03/10/2023
ms.localizationpriority: medium
---

# Manage the sensitivity label of your message or appointment in compose mode (preview)

Collaboration at the workplace not only occurs within the organization, but extends to external partners as well. With information being shared beyond an organization's network, it's important to establish measures to prevent data loss and enforce compliance policies. [Microsoft Purview Information Protection](/microsoft-365/compliance/information-protection) helps you implement solutions to classify and protect sensitive information. The use of sensitivity labels in Outlook is a capability you can configure to protect your data.

You can use the Office JavaScript API to implement sensitivity label solutions in your Outlook add-in projects and support the following scenarios.

- Automatically apply sensitivity labels to certain messages and appointments while they're being composed, so that users can focus on their work.
- Restrict additional actions if a certain sensitivity label is applied to a message or appointment, such as preventing users from adding external recipients to a message.
- Add a header or footer to a message or appointment based on its sensitivity label to comply with business and legal policies.

> [!IMPORTANT]
> Features in preview shouldn't be used in production add-ins. We invite you to test this feature in test or development environments and welcome feedback on your experience through GitHub (see the **Feedback** section at the end of this page).

## Supported clients and platforms

The following table lists client-server combinations that support the use of the sensitivity label feature in Outlook add-ins. Excluded combinations aren't supported.

|Client|Exchange Online|
|------|------|
|**Windows**<br>Version 2302 (Build 16130.20020) or later|Supported|
|**Mac**|Supported|
|**Web browser (modern UI)**|Supported|
|**iOS**|Not applicable|
|**Android**|Not applicable|

## Preview the sensitivity label feature

To test the sensitivity label feature in your add-in while it's in preview, set up your preferred Outlook client accordingly.

- For Outlook on Windows, install Version 2302 (Build 16130.20020) or later. Once installed, join the [Office Insider program](https://insider.office.com/join/windows) and select the **Beta Channel** option to access Office beta builds.
- For Outlook on Mac, install Build <> or later. Once installed, join the [Office Insider program](https://insider.office.com/join/windows) and select the **Beta Channel** option to access Office beta builds.
- For Outlook on the web, ensure that the Targeted release option is set up on your Microsoft 365 tenant. To learn more, see the "Targeted release" section of [Set up the Standard or Targeted release options](/microsoft-365/admin/manage/release-options-in-office-365).

## Configure the manifest

To be able to use the sensitivity feature in your Outlook add-in project, you must set the [\<Permissions\> element](/javascript/api/manifest/permissions) of the XML manifest to **ReadWriteItem**.

```xml
<Permissions>ReadWriteItem</Permissions>
```

If your add-in will detect and handle the `OnSensitivityLabelChanged` event, additional manifest configurations are required to enable the event-based activation feature. To learn more, see [Detect sensitivity label changes with the OnSensitivityLabelChanged event](#detect-sensitivity-label-changes-with-the-onsensitivitylabelchanged-event).

> [!IMPORTANT]
> The sensitivity label feature isn't yet supported for the [Teams manifest for Office Add-ins (preview)](../develop/json-manifest-overview.md).

## Verify the status of the catalog of sensitivity labels

Sensitivity labels and their policies are configured by an organization's administrator through the [Microsoft Purview compliance portal](/microsoft-365/compliance/microsoft-365-compliance-center). For guidance on how to configure sensitivity labels in your tenant, see [Create and configure sensitivity labels and their policies](/microsoft-365/compliance/create-sensitivity-labels).

Before you can get or set the sensitivity label on a message or appointment, you must first ensure that the catalog of sensitivity labels is enabled on the mailbox where the add-in is installed. To check the status of the catalog of sensitivity labels, call [context.sensitivityLabelsCatalog.getIsEnabledAsync](/javascript/api/outlook/office.sensitivitylabelscatalog?view=outlook-js-preview&preserve-view=true#outlook-office-sensitivitylabelscatalog-getisenabledasync-member(1)) in compose mode.

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

If you want to determine the sensitivity labels available for use on a message or appointment in compose mode, use [context.sensitivityLabelsCatalog.getAsync](/javascript/api/outlook/office.sensitivitylabelscatalog?view=outlook-js-preview&preserve-view=true#outlook-office-sensitivitylabelscatalog-getasync-member(1)). The available labels are returned in the form of [SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails?view=outlook-js-preview&preserve-view=true) objects, which provide the following details.

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

To get the sensitivity label currently applied to a message or appointment in compose mode, call [item.sensitivityLabel.getAsync](/javascript/api/outlook/office.sensitivitylabel?view=outlook-js-preview&preserve-view=true#outlook-office-sensitivitylabel-getasync-member(1)) as shown in the following example. This returns the GUID of the sensitivity label currently applied.

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

You can set only one sensitivity label on a message or appointment in compose mode. Before you set the label, call [context.sensitivityLabelsCatalog.getAsync](/javascript/api/outlook/office.sensitivitylabelscatalog?view=outlook-js-preview&preserve-view=true#outlook-office-sensitivitylabelscatalog-getasync-member(1)). This ensures that the label you want to apply is available for use. It also helps you identify a label's GUID, which you'll need to apply the label to the mail item. After you confirm the label's availability, pass its GUID as a parameter to [item.sensitivityLabel.setAsync](/javascript/api/outlook/office.sensitivitylabel?view=outlook-js-preview&preserve-view=true#outlook-office-sensitivitylabel-setasync-member(1)), as shown in the following example.

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

Instead of using the GUID to set the sensitivity label, you can pass the [SensitivityLabelDetails](/javascript/api/outlook/office.sensitivitylabeldetails?view=outlook-js-preview&preserve-view=true) object retrieved from the catalog call, as shown in the following example.

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

The `OnSensitivityLabelChanged` event is available through the event-based activation feature. To learn how to configure, debug, and deploy an event-based add-in that uses this event, see [Configure your Outlook add-in for event-based activation](autolaunch.md).

> [!IMPORTANT]
> The `OnSensitivityLabelChanged` event isn't yet supported for the [Teams manifest for Office Add-ins (preview)](../develop/json-manifest-overview.md).

## See also

- [Learn about sensitivity labels](/microsoft-365/compliance/sensitivity-labels)
- [Get started with sensitivity labels](/microsoft-365/compliance/get-started-with-sensitivity-labels)
- [Create and configure sensitivity labels and their policies](/microsoft-365/compliance/create-sensitivity-labels)
- [Configure your Outlook add-in for event-based activation](autolaunch.md)
