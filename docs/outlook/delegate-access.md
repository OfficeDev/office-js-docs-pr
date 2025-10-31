---
title: Implement shared folders and shared mailbox scenarios in an Outlook add-in
description: Discusses how to configure Outlook add-in support for shared folders (also known as delegate access) and shared mailboxes.
ms.date: 09/11/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement shared folders and shared mailbox scenarios in an Outlook add-in

This article describes how to implement shared folders (also known as delegate access) and shared mailbox scenarios in your Outlook add-in, including which permissions the Office JavaScript API supports.

> [!NOTE]
> Shared folder support was introduced in [requirement set 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8), while shared mailbox support was introduced in [requirement set 1.13](/javascript/api/requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13). For information about client support for these features, see [Supported clients and platforms](#supported-clients-and-platforms).

## Supported clients and platforms

The following table shows supported client-server combinations for this feature, including the minimum required Cumulative Update where applicable.

| Client | Exchange Online | Exchange Server Subscription Edition (SE) | Exchange 2019 on-premises<br>(Cumulative Update 1 or later) | Exchange 2016 on-premises<br>(Cumulative Update 6 or later) |
|---|---|---|---|---|
|**Web browser (modern Outlook UI)**|Supported|Not applicable|Not applicable|Not applicable|
|**Web browser (classic Outlook UI)**|Not applicable|<ul><li>**Shared folders**: Supported</li><li>**Shared mailboxes**: Not applicable</li></ul>|<ul><li>**Shared folders**: Supported</li><li>**Shared mailboxes**: Not applicable</li></ul>|<ul><li>**Shared folders**: Supported</li><li>**Shared mailboxes**: Not applicable</li></ul>|
|[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)|Supported|Not applicable|Not applicable|Not applicable|
|**Windows (classic)**<br>**Shared folders**: Version 1910 (Build 12130.20272) or later<br><br>**Shared mailboxes**: Version 2304 (Build 16327.20248) or later|Supported|Supported\*|Supported\*|Supported\*|
|**Mac**<br>Version 16.47 or later|Supported|Supported|Supported|Supported|
|**Android**|Not applicable|Not applicable|Not applicable|Not applicable|
|**iOS**|Not applicable|Not applicable|Not applicable|Not applicable|

> [!NOTE]
> \* Support for this feature in an on-premises Exchange environment is available starting in classic Outlook on Windows Version 2206 (Build 15330.20000) for the Current Channel and Version 2207 (Build 15427.20000) for the Monthly Enterprise Channel.

## Supported setups

The following sections describe configurations for shared mailboxes and shared folders that support the use of add-ins. The feature APIs may not work as expected in other configurations. Select the platform you'd like to learn how to configure.

### [Web (modern) and new Outlook on Windows](#tab/web)

#### Shared folders

The mailbox owner must first provide access to a delegate.

- To provide access to manage meetings and meeting responses on behalf of the mailbox owner, see [Calendar delegation in Outlook on the web](https://support.microsoft.com/office/532e6410-ee80-42b5-9b1b-a09345ccef1b).

- To provide access to manage both the inbox and calendar on behalf of the mailbox owner, access must be configured through one of the following options.

  - The mailbox owner can configure access through classic Outlook on Windows. To learn more, see [Allow someone else to manage your mail and calendar](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926).

  - An administrator can configure access through the Microsoft 365 admin center. To learn more, see [Give mailbox permissions to another Microsoft 365 user](/microsoft-365/admin/add-users/give-mailbox-permissions-to-another-user).

  - An administrator can configure access through the Exchange admin center. To learn more, see [Manage permissions for recipients](/exchange/recipients/mailbox-permissions).

Once access is provided, the delegate must then follow the instructions outlined in [Access another person's mailbox](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081).

In new Outlook on Windows, by default, shared mailboxes that are automatically mapped by an administrator are added as shared folders. This means that while a user can read and send messages from the shared mailbox, they can't manage the mailbox settings. To manage the settings, a user must promote the shared mailbox to a full account. For more information, see [Manage shared mailbox settings in new Outlook](https://support.microsoft.com/office/f6929a97-4fc6-4a52-b77d-5e596c6322b4).

#### Shared mailboxes

A shared mailbox allows a group of users to easily monitor and send messages and meeting invites using a shared email address.

In Outlook on the web, a shared mailbox can be opened in the same panel as a user's primary mailbox or in a separate browser tab or window. For guidance, see [Open and use a shared mailbox in Outlook on the web](https://support.microsoft.com/office/98b5a90d-4e38-415d-a030-f09a4cd28207).

In new Outlook on Windows, a shared mailbox is accessed from the client's folder pane. A shared mailbox can be automatically added by an administrator or manually added by the user. Mailboxes that are manually added are automatically set up as full accounts in the Outlook client, so that users can manage the mailbox settings. Conversely, by default, shared mailboxes added by an administrator are set up as shared folders. If a user wants to manage the settings of the mailbox, they must promote the shared folder to a full account on the client. For more information, see [Manage shared mailbox settings in new Outlook](https://support.microsoft.com/office/f6929a97-4fc6-4a52-b77d-5e596c6322b4).

### [Windows (classic)](#tab/windows)

#### Shared folders

The mailbox owner must first provide access to a delegate using one of the following options.

- Set up delegate access from the mailbox in classic Outlook on Windows. To learn more, see [Allow someone else to manage your mail and calendar](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926).

- Set up delegate access from the Microsoft 365 admin center. This option can only be completed by administrators. To learn more, see [Give mailbox permissions to another Microsoft 365 user](/microsoft-365/admin/add-users/give-mailbox-permissions-to-another-user).

- Set up delegate access from the Exchange admin center. This option can only be completed by administrators. To learn more, see [Manage permissions for recipients](/exchange/recipients/mailbox-permissions).

Once access is provided, the delegate must then follow the instructions outlined in the "Add another person's mailbox to your profile" section of the article [Manage another person's mail and calendar items](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5).

#### Shared mailboxes

Exchange server admins can create and manage shared mailboxes for sets of users to access. [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) and [on-premises Exchange environments](/exchange/collaboration/shared-mailboxes/create-shared-mailboxes) are supported.

An Exchange Server feature known as "automapping" is on by default. This means that subsequently the [shared mailbox should automatically appear](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) in a user's Outlook client after Outlook has been closed and reopened. However, if an admin turned off automapping, the user must follow the manual steps outlined in the "Add a shared mailbox to Outlook" section of the article [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> Do **NOT** sign into the shared mailbox with a password. The feature APIs won't work in that case.

### [Mac](#tab/unix)

#### Shared folders

The mailbox owner must first provide access to a delegate. For guidance, see [Add and manage delegates in Outlook for Mac](https://support.microsoft.com/office/49ba7631-1984-453e-8a8f-c78fd43475e4).

Once access is provided, the delegate must then follow the instructions outlined in [Become a delegate or stop being a delegate in Outlook for Mac](https://support.microsoft.com/office/818da19c-b03c-4a69-926a-76a7c84c3579).

#### Shared mailboxes

A shared mailbox allows a group of users to easily monitor and send messages and meeting invites using a shared email address. For guidance on how to access a shared mailbox that you have permissions to in Outlook on Mac, see the "Open a shared or delegated mailbox" section of [Open a shared Mail, Calendar, or People folder in Outlook for Mac](https://support.microsoft.com/office/6ecc39c5-5577-4a1d-b18c-bbdc92972cb2).

---

## Configure the manifest

To implement shared folder and shared mailbox scenarios in your add-in, you must first configure support for the feature in your manifest. The markup varies depending on the type of manifest your add-in uses.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

> [!NOTE]
> Add-ins that use the unified manifest for Microsoft 365 aren't directly supported in Outlook on Mac. To run this type of add-in in Outlook on Mac, the add-in must first be published to [Microsoft Marketplace](https://marketplace.microsoft.com/) then deployed in the [Microsoft 365 Admin Center](../publish/publish.md). For more information, see [Support for add-ins with the unified manifest for Microsoft 365](compare-outlook-add-in-support-in-outlook-for-mac.md#support-for-add-ins-with-the-unified-manifest-for-microsoft-365).

Add an additional object to the [`"authorization.permissions.resourceSpecific"`](/microsoft-365/extensibility/schema/root-authorization-permissions#resourcespecific) array. Set its `"name"` property to `"Mailbox.SharedFolder"` and its `"type"` property to `"Delegated"`.

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

Under the parent `<DesktopFormFactor>` element, set the [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) element to `true`. At present, other form factors aren't supported.

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

> [!NOTE]
> In Outlook on the web and on Windows (new and classic), depending on how the shared folder or mailbox is accessed, the `getSharedPropertiesAsync` method may require certain conditions to be met in Message Compose mode. For more information, see the "Message Compose mode" section in [Limitations](#message-compose-mode).

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
>
> - Exchange Web Services (EWS) isn't supported in shared folder and shared mailbox scenarios.
>
> - In delegate or shared scenarios, a delegate can get the [categories from the Outlook master list](/javascript/api/outlook/office.mastercategories) but can't add or remove categories from the list.

### Configure the add-in's permissions

To use Microsoft Graph services, an add-in must configure the **read/write mailbox** permission in its manifest. The markup varies depending on the type of manifest your add-in uses.

- **Unified manifest for Microsoft 365**: Set the `"name"` property of an object in the `"authorization.permissions.resourceSpecific"` array to `"Mailbox.ReadWrite.User"`.
- **Add-in only manifest**: Set the [Permissions](/javascript/api/manifest/permissions) element to **ReadWriteMailbox**.

### Use Microsoft Graph

To implement your shared folder and shared mailbox scenarios, use Microsoft Graph to access additional mailbox information and resources. For example, you can use Microsoft Graph to [get the contents of an Outlook message that's attached to a message](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post) in a mailbox where a user has delegate access. For guidance on how to use Microsoft Graph, see [Overview of Microsoft Graph](/graph/overview) and [Outlook mail API in Microsoft Graph](/graph/outlook-mail-concept-overview).

> [!TIP]
> To access Microsoft Graph APIs from your add-in, use MSAL.js nested app authentication (NAA). To learn more, see [Enable SSO in an Office Add-in using nested app authentication (preview)](../develop/enable-nested-app-authentication-in-your-add-in.md).

## Add-in support in shared folder and shared mailbox scenarios

The availability of add-ins in a shared folder or shared mailbox varies depending on the scenario and Outlook client.

> [!TIP]
> To learn more about where add-ins do and don't activate in general, see the [Add-in activation limitations](outlook-add-ins-overview.md#add-in-activation-limitations) section of the Outlook add-ins overview page.

### Add-ins in shared folder scenarios

The following table outlines the availability of add-ins in shared folder scenarios.

| Scenario | Add-in availability |
| ----- | ----- |
| **Inbox** folder is shared with a delegate | Add-ins are available to the delegate in message read mode. |
| **Drafts** folder is shared with a delegate | Add-ins are available to the delegate in message compose mode. |
| (New Outlook on Windows only) Shared mailbox is automatically mapped by an administrator and isn't promoted to a full account by the user | See the behaviors outlined in the "Web: same tab or window, Windows (new): non-promoted mailbox, Windows (classic), and Mac" column of [Add-ins in shared mailbox scenarios](#add-ins-in-shared-mailbox-scenarios). |
| Calendar is explicitly shared with a delegate (the entire mailbox may not be shared) | Add-ins are available to the delegate in appointment read and compose modes. |
| Calendar is shared with a group of users with different access (for example, made it editable to a particular distribution list or the entire organization) | Add-ins are available to users with indirect or implicit permissions in appointment read and compose modes. |

### Add-ins in shared mailbox scenarios

The following table outlines the availability of add-ins in shared mailbox scenarios across various Outlook clients. Note that the behavior in Outlook on the web may differ depending on whether the shared mailbox is opened in the same panel as the user's primary mailbox or in a separate tab or window using the **Open another mailbox** option. Similarly, the behavior in new Outlook on Windows may also differ depending on whether the shared mailbox was added or promoted as a full account on the client.

| Scenario | Applicable Outlook clients<ul><li>Web: same tab or window</li><li>Windows (new): non-promoted</li><li>Windows (classic)</li><li>Mac</li></ul> | Applicable Outlook clients<ul><li>Web: separate tab or window</li><li>Windows (new): promoted</li></ul> |
| ----- | ----- | ----- |
| Add-in installed by the user | Users can't install add-ins in a shared mailbox. Add-ins installed by a user are added to the user's primary mailbox. | Users can't install add-ins in a shared mailbox. The in-app Microsoft 365 and Copilot store doesn't appear on the ribbon of the mailbox. |
| Add-in installed by an administrator | Administrators shouldn't deploy add-ins to a shared mailbox. They should instead deploy an add-in to the user's primary mailbox. The user can then use the add-in in a shared mailbox as long as the add-in meets certain requirements (see the following scenarios for add-in availability in read and compose modes). | The same limitation and recommendation on other platforms apply (see previous column). |
| Add-in used in read mode | An add-in's manifest must be configured to support shared mailbox scenarios. For more information, see [Configure the manifest](#configure-the-manifest). The add-in must be installed in the user's primary mailbox by the user or administrator. | The same manifest configuration and behavior on other platforms apply (see previous column). |
| Add-in used in compose mode | In Outlook on the web (mailbox opened in the same window) and on Windows (new and classic), add-ins installed in the user's primary mailbox that support compose mode are available for use. An add-in's manifest doesn't need additional configuration to support shared mailbox scenarios.<br><br>However, in Outlook on Mac, an add-in's manifest must be configured to support shared mailbox scenarios. For more information, see [Configure the manifest](#configure-the-manifest). | An add-in's manifest must be configured to support shared mailbox scenarios. For more information, see [Configure the manifest](#configure-the-manifest). The add-in must be installed in the user's primary mailbox by the user or administrator. |
| Templates created using the My Templates add-in | This only applies to Outlook on the web and on Windows (new and classic) since the My Templates add-in isn't supported in shared mailboxes on Outlook on Mac.<br><br>Templates created are saved to the creator's primary mailbox. Although the creator can use these templates in both their primary and shared mailboxes, other users who have access to the shared mailbox can't access these templates. For more information, see [Create an email message template](https://support.microsoft.com/office/43ec7142-4dd0-4351-8727-bd0977b6b2d1). | Templates created are saved to the shared mailbox. Anyone with access to the shared mailbox can edit or use these templates if they open the mailbox using **Open another mailbox** in Outlook on the web, or if the mailbox is promoted to a full account in new Outlook on Windows. These shared templates can't be accessed by anyone, including the template creator, from other platforms. This includes Outlook on the web, if the shared mailbox is opened in the same tab as the user's primary mailbox, and new Outlook on Windows, if the shared mailbox wasn't promoted to a full account. Conversely, templates created on other platforms can't be accessed from a shared mailbox opened using the **Open another mailbox** option or from a promoted shared mailbox. For more information, see [Create an email message template](https://support.microsoft.com/office/43ec7142-4dd0-4351-8727-bd0977b6b2d1). |
| Default add-ins in Outlook | In Outlook on the web and on Windows (new and classic), default Outlook add-ins are available for use in a shared mailbox. Default Outlook add-ins can include My Templates, Unsubscribe, and Action Items. Note that some default add-ins may not appear in your organization.<br><br>In Outlook on Mac, default add-ins aren't available in a shared mailbox. | Default add-ins are available in a shared mailbox. |

## Limitations

Depending on your add-in's scenarios, there are a few limitations for you to consider when handling shared folder or shared mailbox situations.

### Message Compose mode

In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) isn't supported in Outlook on the web or on Windows (new and classic) unless the following conditions are met.

- **Delegate access/Shared folders**

    1. The mailbox owner starts a message. This can be a new message, a reply, or a forward.
    1. They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.
    1. The delegate opens the draft from the shared folder then continues composing.

- **Shared mailbox opened in the same panel as the user's primary mailbox (web, classic Windows) or shared mailbox that hasn't been promoted to a full account (new Windows)**

    1. A shared mailbox user starts a message. This can be a new message, a reply, or a forward.
    1. They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.
    1. Another shared mailbox user opens the draft from the shared mailbox then continues composing.

  > [!NOTE]
  > The `getSharedPropertiesAsync` method is supported on the following platforms without additional conditions.
  >
  > - Outlook on the web when the shared mailbox is opened in a separate tab or window using the **Open another mailbox** option.
  > - new Outlook on Windows when the shared mailbox is promoted to a full account.

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
