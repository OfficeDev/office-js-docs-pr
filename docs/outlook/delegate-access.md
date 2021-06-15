---
title: Enable shared folders and shared mailbox scenarios in an Outlook add-in
description: 'Discusses how to configure add-in support for shared folders and shared mailboxes.'
ms.date: 06/15/2021
localization_priority: Normal
---

# Enable shared folders and shared mailbox scenarios in an Outlook add-in

This article describes how to enable shared folders (also known as delegate access) scenarios in your Outlook add-in, including which permissions the Office JavaScript API supports. Shared mailbox support is now in preview.

> [!IMPORTANT]
> Support for this feature was introduced in requirement set 1.8. See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Supported setups

The following sections describe supported configurations for shared mailboxes (now in preview) and shared folders. The feature APIs may not work as expected in other configurations.

To learn more about where add-ins do and do not activate in general, refer to the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page.

### Shared folders

After the mailbox owner has provided access to a delegate, the delegate can then [access and manage that person's mailbox](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5). The delegate must follow the instructions outlined in the "Add another person's mailbox to your profile" section of that article.

### Shared mailboxes (preview)

Exchange server admins can create and manage shared mailboxes on [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) for sets of users to access.

An Exchange Server feature known as "automapping" is on by default which means that subsequently the [shared mailbox should automatically appear](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) in a user's Outlook app after Outlook has been closed and reopened. However, if an admin turned off automapping, the user must follow the manual steps outlined in the "Add a shared mailbox to Outlook" section of the article [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> Do **NOT** sign into the shared mailbox with a password. The feature APIs won't work in that case.

## Supported permissions

The following table describes the permissions that the Office JavaScript API supports for delegates and shared mailbox users.

|Permission|Value|Description|
|---|---:|---|
|Read|1 (000001)|Can read items.|
|Write|2 (000010)|Can create items.|
|DeleteOwn|4 (000100)|Can delete only the items they created.|
|DeleteAll|8 (001000)|Can delete any items.|
|EditOwn|16 (010000)|Can edit only the items they created.|
|EditAll|32 (100000)|Can edit any items.|

> [!NOTE]
> Currently the API supports getting existing permissions, but not setting permissions.

The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the permissions. Each position in the bitmask represents a particular permission and if it's set to `1` then the user has the respective permission. For example, if the second bit from the right is `1`, then the user has **Write** permission. You can see an example of how to check for a specific permission in the [Perform an operation as delegate or shared mailbox user](#perform-an-operation-as-delegate-or-shared-mailbox-user) section later in this article.

## Sync across shared folder clients

A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.

However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay. To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.

> [!IMPORTANT]
> In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.

## Configure the manifest

To enable shared folders and shared mailbox scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`. At present, other form factors are not supported.

To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.

The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.

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
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## Perform an operation as delegate or shared mailbox user

You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method. This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the user's permissions, the owner's email address, the REST API's base URL, and the target mailbox.

The following example shows how to get the shared properties of a message or appointment, check if the delegate or shared mailbox user has **Write** permission, and make a REST call.

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).

## Handle calling REST on shared and non-shared items

If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared. After that, you can construct the REST URL for the operation using the appropriate object.

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## Limitations in shared folder scenarios

Depending on your add-in's scenarios, there are a few limitations for you to consider when handling delegate situations.

### Message Compose mode

In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or on Windows unless the following conditions are met.

1. The owner shares at least one mailbox folder with the delegate.
1. The delegate drafts a message in the shared folder.

    Examples:

    - The delegate replies to or forwards an email in the shared folder.
    - The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder. The delegate opens the draft from the shared folder then continues composing.

After the message has been sent, it's usually found in the delegate's **Sent Items** folder.

### REST and EWS

Your add-in can use REST and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox. EWS is not supported.

## See also

- [Allow someone else to manage your mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Calendar sharing in Microsoft 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Add a shared mailbox to Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [How to order manifest elements](../develop/manifest-element-ordering.md)
- [Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript bitwise operators](https://www.w3schools.com/js/js_bitwise.asp)