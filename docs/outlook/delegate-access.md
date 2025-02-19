---
title: Implement shared folders and shared mailbox scenarios in an Outlook add-in
description: Discusses how to configure Outlook add-in support for shared folders (also known as delegate access) and shared mailboxes.
ms.date: 02/20/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement shared folders and shared mailbox scenarios in an Outlook add-in

This article describes how to implement shared folders (also known as delegate access) and shared mailbox scenarios in your Outlook add-in, including which permissions the Office JavaScript API supports.

> [!NOTE]
> Shared folder support was introduced in [requirement set 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8), while shared mailbox support was introduced in [requirement set 1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13). For information about client support for these features, see [Supported clients and platforms](#supported-clients-and-platforms).

## Supported clients and platforms

The following table shows supported client-server combinations for this feature, including the minimum required Cumulative Update where applicable. Excluded combinations aren't supported.

| Client | Exchange Online | Exchange 2019 on-premises<br>(Cumulative Update 1 or later) | Exchange 2016 on-premises<br>(Cumulative Update 6 or later) |
|---|---|---|---|
|**Web browser (modern Outlook UI)**|Supported<sup>1</sup>|Not applicable|Not applicable|
|**Web browser (classic Outlook UI)**|Not applicable|<ul><li>**Shared folders**: Supported</li><li>**Shared mailboxes**: Not applicable</li></ul>|<ul><li>**Shared folders**: Supported</li><li>**Shared mailboxes**: Not applicable</li></ul>|
|[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)|Supported|Not applicable|Not applicable|
|**Windows (classic)**<br>**Shared folders**: Version 1910 (Build 12130.20272) or later<br><br>**Shared mailboxes**: Version 2304 (Build 16327.20248) or later|Supported|Supported<sup>2</sup>|Supported<sup>2</sup>|
|**Mac**<br>Version 16.47 or later|Supported|Supported|Supported|

> [!NOTE]
> <sup>1</sup>In Outlook on the web, if you open a shared mailbox in a separate browser tab or window using the **Open another mailbox** option, you may encounter issues when accessing add-ins from the mailbox. We recommend opening the mailbox in the same panel as your primary mailbox instead. This ensures that add-ins work as expected in your shared mailbox. If you prefer to open the shared mailbox using the **Open another mailbox** option, we recommend deploying the add-in to both your primary user and shared mailboxes.
>
> <sup>2</sup>Support for this feature in an on-premises Exchange environment is available starting in classic Outlook on Windows Version 2206 (Build 15330.20000) for the Current Channel and Version 2207 (Build 15427.20000) for the Monthly Enterprise Channel.

## Supported setups

The following sections describe supported configurations for shared mailboxes and shared folders. The feature APIs may not work as expected in other configurations. Select the platform you'd like to learn how to configure.

### [Web (modern) and new Outlook on Windows](#tab/web)

#### Shared folders

The mailbox owner must first provide access to a delegate.

- To provide access to manage meetings and meeting responses on behalf of the mailbox owner, see [Calendar delegation in Outlook on the web](https://support.microsoft.com/office/532e6410-ee80-42b5-9b1b-a09345ccef1b).

- To provide access to manage both the inbox and calendar on behalf of the mailbox owner, access must be configured through one of the following options.

  - The mailbox owner can configure access through classic Outlook on Windows. To learn more, see [Allow someone else to manage your mail and calendar](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926).

  - An administrator can configure access through the Microsoft 365 admin center. To learn more, see [Give mailbox permissions to another Microsoft 365 user](/microsoft-365/admin/add-users/give-mailbox-permissions-to-another-user).

  - An administrator can configure access through the Exchange admin center. To learn more, see [Manage permissions for recipients](/exchange/recipients/mailbox-permissions).

Once access is provided, the delegate must then follow the instructions outlined in the "Add another person’s mailbox to your folder list in Outlook Web App" section of the article [Access another person's mailbox](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081).

#### Shared mailboxes

A shared mailbox allows a group of users to easily monitor and send messages and meeting invites using a shared email address.

In Outlook on the web, a shared mailbox can be opened in the same panel as a user's primary mailbox or in a separate browser tab or window. For guidance, see [Open and use a shared mailbox in Outlook on the web](https://support.microsoft.com/office/98b5a90d-4e38-415d-a030-f09a4cd28207).

> [!NOTE]
> In Outlook on the web, if you open a shared mailbox in a separate browser tab or window using the **Open another mailbox** option, you may encounter issues when accessing add-ins from the mailbox. We recommend opening the mailbox in the same panel as your primary mailbox instead. This ensures that add-ins work as expected in your shared mailbox.
>
> If you prefer to open the shared mailbox using the **Open another mailbox** option, we recommend deploying the add-in to both your primary user and shared mailboxes.

In new Outlook on Windows, a shared mailbox is added to the **Shared with me** section of the folder pane. For guidance, see [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd).

### [Windows (classic)](#tab/windows)

#### Shared folders

The mailbox owner must first provide access to a delegate using one of the following options.

- Set up delegate access from the mailbox in classic Outlook on Windows. To learn more, see [Allow someone else to manage your mail and calendar](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926).

- Set up delegate access from the Microsoft 365 admin center. This option can only be completed by administrators. To learn more, see [Give mailbox permissions to another Microsoft 365 user](/microsoft-365/admin/add-users/give-mailbox-permissions-to-another-user).

- Set up delegate access from the Exchange admin center. This option can only be completed by administrators. To learn more, see [Manage permissions for recipients](/exchange/recipients/mailbox-permissions).

Once access is provided, the delegate must then follow the instructions outlined in the "Add another person's mailbox to your profile" section of the article [Manage another person's mail and calendar items](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5).

#### Shared mailboxes

Exchange server admins can create and manage shared mailboxes for sets of users to access. [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) and [on-premises Exchange environments](/exchange/collaboration/shared-mailboxes/create-shared-mailboxes) are supported.

An Exchange Server feature known as "automapping" is on by default which means that subsequently the [shared mailbox should automatically appear](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) in a user's Outlook app after Outlook has been closed and reopened. However, if an admin turned off automapping, the user must follow the manual steps outlined in the "Add a shared mailbox to Outlook" section of the article [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> Do **NOT** sign into the shared mailbox with a password. The feature APIs won't work in that case.

### [Mac](#tab/unix)

#### Shared mailboxes

A shared mailbox allows a group of users to easily monitor and send messages and meeting invites using a shared email address. For guidance on how to access a shared mailbox that you have permissions to in Outlook on Mac, see the "Open a shared or delegated mailbox" section of [Open a shared Mail, Calendar, or People folder in Outlook for Mac](https://support.microsoft.com/office/6ecc39c5-5577-4a1d-b18c-bbdc92972cb2).

Users with permissions to a shared mailbox can activate add-ins configured for shared mailbox scenarios in message and appointment read and compose modes.

#### Shared folders

If the **Inbox** folder is shared with a delegate, add-ins are available to the delegate in message read mode.

If the **Drafts** folder is also shared with the delegate, add-ins are available in compose mode.

#### Local shared calendar (new model)

If the calendar owner explicitly shared their calendar with a delegate (the entire mailbox may not be shared), add-ins are available to the delegate in appointment read and compose modes.

#### Remote shared calendar (previous model)

If the calendar owner granted broad access to their calendar (for example, made it editable to a particular DL or the entire organization), users may then have indirect or implicit permission and add-ins are available to those users in appointment read and compose modes.

---

To learn more about where add-ins do and don't activate in general, refer to the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page.

## Configure the manifest

To implement shared folders and shared mailbox scenarios in your add-in, you must first configure support for the feature in your manifest. The markup varies depending on the type of manifest your add-in uses.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

Add an additional object to the "authorization.permissions.resourceSpecific" array and set its "name" property to "Mailbox.SharedFolder".

```json
"authorization": {
  "permissions": {
    "resourceSpecific": [
      ...
      {
        "name": "Mailbox.SharedFolder",
        "type": "Delegated"
      },
    ]
  }
},
```

# [Add-in only manifest](#tab/xmlmanifest)

Under the parent **\<DesktopFormFactor\>** element, set the [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) element to `true`. At present, other form factors aren't supported.

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure the extension point. -->
          </ExtensionPoint>
          ...
        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

---

## Identify if a folder or mailbox is shared

Before you can run operations in a shared folder or shared mailbox, you must first identify whether the current folder or mailbox is shared. To determine this, call [Office.context.mailbox.item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) on a message or appointment in compose or read mode. If the item is in a shared folder or shared mailbox, the method returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that provides the user's permissions, the owner's email address, the REST API's base URL, and the location of the target mailbox.

The following example calls the `getSharedPropertiesAsync` method to identify the owner of the mailbox and the permissions of the delegate or shared mailbox user.

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync((result) => {
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.error("The current folder or mailbox isn't shared.");
    return;
  }
  const sharedProperties = result.value;
  console.log(`Owner: ${sharedProperties.owner}`);
  console.log(`Permissions: ${sharedProperties.delegatePermissions} `);
});
```

### Supported permissions

The following table describes the permissions that `getSharedPropertiesAsync` supports for delegates and shared mailbox users.

|Permission|Value|Description|
|---|---:|---|
|Read|1 (000001)|Can read items.|
|Write|2 (000010)|Can create items.|
|DeleteOwn|4 (000100)|Can delete only the items they created.|
|DeleteAll|8 (001000)|Can delete any items.|
|EditOwn|16 (010000)|Can edit only the items they created.|
|EditAll|32 (100000)|Can edit any items.|

> [!NOTE]
> Currently, the API supports getting existing permissions, but not setting permissions.

The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) enum returned by the [delegatePermissions](/javascript/api/outlook/office.sharedproperties#outlook-office-sharedproperties-delegatepermissions-member) property is implemented using a bitmask to indicate the permissions. Each position in the bitmask represents a particular permission and if it's set to `1`, then the user has the respective permission. For example, if the second bit from the right is `1`, then the user has **Write** permission.

## Perform an operation as a delegate or shared mailbox user

Once you've identified that the current mail item is in a shared folder or shared mailbox, your add-in can then perform the necessary operations on the item within the shared environment. To run operations on an item in a shared context, you must first configure your add-in's permission in the manifest. Then, use Microsoft Graph to complete the operations.

> [!NOTE]
> Exchange Web Services (EWS) isn't supported in shared folder and shared mailbox scenarios.

### Configure the add-in's permissions

To use Microsoft Graph services, an add-in must configure the **read/write mailbox** permission in its manifest. The markup varies depending on the type of manifest your add-in uses.

- **Unified manifest for Microsoft 365**: Set the "name" property of an object in the "authorization.permissions.resourceSpecific" array to "Mailbox.ReadWrite.User".
- **Add-in only manifest**: Set the [Permissions](/javascript/api/manifest/permissions) element to **ReadWriteMailbox**.

### Use Microsoft Graph

To implement your shared folder and shared mailbox scenarios, use Microsoft Graph to access additional mailbox information and resources. For example, you can use Microsoft Graph to [get the contents of an Outlook message that's attached to a message](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post) in a mailbox where a user has delegate access. For guidance on how to use Microsoft Graph, see [Overview of Microsoft Graph](/graph/overview) and [Outlook mail API in Microsoft Graph](/graph/outlook-mail-concept-overview).

> [!TIP]
> To access Microsoft Graph APIs from your add-in, use MSAL.js nested app authentication (NAA). To learn more, see [Enable SSO in an Office Add-in using nested app authentication (preview)](../develop/enable-nested-app-authentication-in-your-add-in.md).

## Limitations

Depending on your add-in's scenarios, there are a few limitations for you to consider when handling shared folder or shared mailbox situations.

### Message Compose mode

In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) isn't supported in Outlook on the web or on Windows (new and classic) unless the following conditions are met.

- **Delegate access/Shared folders**

    1. The mailbox owner starts a message. This can be a new message, a reply, or a forward.
    1. They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.
    1. The delegate opens the draft from the shared folder then continues composing.

- **Shared mailbox (applies to classic Outlook on Windows only)**

    1. A shared mailbox user starts a message. This can be a new message, a reply, or a forward.
    1. They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.
    1. Another shared mailbox user opens the draft from the shared mailbox then continues composing.

Once these conditions are met, the message becomes available in a shared context and add-ins that support these shared scenarios can get the item's shared properties. After the message is sent, it's usually found in the **Sent Items** folder of the sender's personal mailbox.

### User or shared mailbox hidden from an address list

If an admin hid a user or shared mailbox address from an address list, such as the global address list (GAL), affected mail items opened in the mailbox report `Office.context.mailbox.item` as null. For example, if the user opens a mail item in a shared mailbox that's hidden from the GAL, `Office.context.mailbox.item` representing that mail item is null.

### Sync across shared folder clients

A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately. However, if Microsoft Graph operations were used to set an extended property on an item, such changes could take some time to sync. To avoid a delay, we recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs. To learn more, see the "Custom properties" tab of [Get and set metadata in an Outlook add-in](metadata-for-an-outlook-add-in.md?tabs=custom-properties).

## See also

- [Allow someone else to manage your mail and calendar](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [Calendar sharing in Microsoft 365](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Add a shared mailbox to Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [Overview of Microsoft Graph](/graph/overview)
