---
title: Manifest configuration for Outlook Add-ins
description: Get an overview of the add-in manifest markup and JSON that is relevant only to Outlook.
ms.date: 05/24/2023
ms.localizationpriority: high
---

# Manifest configuration for Outlook Add-ins

An Outlook add-in consists of two components: the add-in manifest and a web app supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients.

You can learn about manifests at [Office Add-in manifests](../develop/add-in-manifests.md). This article focuses on aspects of the manifest that are primarily relevant to Outlook.

## Permissions

Outlook add-ins have a special set of permissions that don't apply to add-ins for other Office host applications. These permissions must be configured in the manifest. For details about these permissions and their effects, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).

# [XML Manifest](#tab/xmlmanifest)

The **\<Permissions\>** element contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but doesn't write to item properties like [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), and doesn't call [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) to access any Exchange Web Services operations should specify **ReadItem** permission. The following is an example of setting an Outlook permission.

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```


# [Unified manifest for Microsoft 365 (developer preview)](#tab/jsonmanifest)

The "authorization.permissions.resourceSpecific" property contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but doesn't write to item properties like [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), and doesn't call [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) to access any Exchange Web Services operations should specify "MailboxItem.Read.User" permission. The following is an example of setting an Outlook permission.

```json
 "authorization": {
        "permissions": {
            "resourceSpecific": [
                {
                    "name": "MailboxItem.ReadWrite.User",
                    "type": "Delegated"
                }
            ]
        }
    },
```

---

## Activation rules

> [!NOTE]
> Activation rules aren't supported in Outlook add-ins that use the unified manifest for Microsoft 365.

Activation rules are specified in the **\<Rule\>** element. The **\<Rule\>** element can appear as a child of the **\<OfficeApp\>** element in 1.1 manifests.

Activation rules can be used to activate an add-in based on one or more of the following conditions on the currently selected item.

> [!NOTE]
> Activation rules only apply to clients that don't support the **\<VersionOverrides\>** element.

- The item type and/or message class

- The presence of a specific type of known entity, such as an address or phone number

- A regular expression match in the body, subject, or sender email address

- The presence of an attachment

For details and samples of activation rules, see [Activation rules for Outlook add-ins](activation-rules.md).

## Form settings

> [!NOTE]
> Form settings aren't supported in Outlook add-ins that use the unified manifest for Microsoft 365.

The **\<FormSettings\>** element is used by older Outlook clients, which only support schema 1.1 and not **\<VersionOverrides\>**. Using this element, developers define how the add-in will appear in such clients. There are two parts - **ItemRead** and **ItemEdit**. **ItemRead** is used to specify how the add-in appears when the user reads messages and appointments. **ItemEdit** describes how the add-in appears while the user is composing a reply, new message, new appointment or editing an appointment where they are the organizer.

These settings are directly related to the activation rules in the **\<Rule\>** element. For example, if an add-in specifies that it should appear on a message in compose mode, an **ItemEdit** form must be specified.

## Next steps: add-in commands

After defining a basic manifest, define add-in commands for your add-in. Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way. For more information, see [Add-in commands](../design/add-in-commands.md).

For an example add-in that defines add-in commands, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).

## Next steps: Add mobile support

Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and on Mac. For more information, see [Add support for add-in commands in Outlook on mobile devices](add-mobile-support.md).

## See also

- [Localization for Office Add-ins](../develop/localization.md)
- [Privacy, permissions, and security for Outlook add-ins](privacy-and-security.md)
- [Outlook add-in APIs](apis.md)
- [Office Add-ins manifest](../develop/add-in-manifests.md)
- [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md)
- [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md)
