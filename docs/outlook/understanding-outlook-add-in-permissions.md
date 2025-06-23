---
title: Understanding Outlook add-in permissions
description: Outlook add-ins specify the required permission level in their manifest, which include restricted, read item, read/write item, or read/write mailbox. 
ms.date: 10/17/2024
ms.localizationpriority: medium
---

# Understanding Outlook add-in permissions

Outlook add-ins specify the required permission level in their manifest. There are four available levels.

|Permission level</br>canonical name|add-in only manifest name|unified manifest for Microsoft 365 name|Summary description|
|:-----|:-----|:-----|:-----|
|**restricted**|Restricted|MailboxItem.Restricted.User|Allows access to properties and methods that don't pertain to specific information about the user or mail item.|
|**read item**|ReadItem|MailboxItem.Read.User|In addition to what is allowed in **restricted**, it allows:<ul><li>regular expressions</li><li>Outlook add-in API read access</li><li>getting the item properties and the callback token</li><li>writing custom properties</li></ul>|
|**read/write item**|ReadWriteItem|MailboxItem.ReadWrite.User|In addition to what is allowed in **read item**, it allows:<ul><li>full Outlook add-in API access except `makeEwsRequestAsync`</li><li>setting the item properties</li></ul>|
|**read/write mailbox**|ReadWriteMailbox|Mailbox.ReadWrite.User|In addition to what is allowed in **read/write item**, it allows:<ul><li>creating, reading, writing items and folders</li><li>sending items</li><li>calling [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)</li></ul>|

Permissions are declared in the manifest. The markup varies depending on the type of manifest.

- **Add-in only manifest**:  Use the **\<Permissions\>** element.
- **Unified manifest for Microsoft 365**: Use the `"name"` property of an object in the [`"authorization.permissions.resourceSpecific"`](/microsoft-365/extensibility/schema/root-authorization-permissions#resourcespecific) array.

> [!NOTE]
>
> - There's a supplementary permission needed for add-ins that use the append-on-send feature. With the add-in only manifest, specify the permission in the [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) element. For details, see [Implement append-on-send in your Outlook add-in](../outlook/append-on-send.md). With the unified manifest, specify this permission with the name **Mailbox.AppendOnSend.User** in an additional object in the `"authorization.permissions.resourceSpecific"` array.
> - There's a supplementary permission needed for add-ins that use shared folders. With the add-in only manifest, specify the permission by setting the [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) element to `true`. For details, see [Implement shared folders and shared mailbox scenarios in an Outlook add-in](../outlook/delegate-access.md). With the unified manifest, specify this permission with the name **Mailbox.SharedFolder** in an additional object in the `"authorization.permissions.resourceSpecific"` array.

The four levels of permissions are cumulative: the **read/write mailbox** permission includes the permissions of **read/write item**, **read item** and **restricted**, **read/write item** includes **read item** and **restricted**, and the **read item** permission includes **restricted**.

You can see the permissions requested by a mail add-in before installing it from [AppSource](https://appsource.microsoft.com). You can also see the required permissions of installed add-ins in the Exchange Admin Center.

> [!TIP]
> To make sure that your Outlook add-in specifies the correct permission level, verify the minimum permission level required by each API implemented by your add-in. For information on minimum permission levels, see [Outlook API reference](/javascript/api/outlook).

## restricted permission

The **restricted** permission is the most basic level of permission. Outlook assigns this permission to a mail add-in by default if the add-in doesn't request a specific permission in its manifest.

### Can do

Access any properties and methods that do **not** pertain to specific information about the user or item (see the next section for the list of members that do).

### Can't do

- Use the [ItemHasAttachment](/javascript/api/manifest/rule#itemhasattachment-rule) or [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) rule.

- Access the members in the following list that pertain to the information of the user or item. Attempting to access members in this list will return **null** and result in an error message which states that Outlook requires the mail add-in to have elevated permission.

  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.userProfile](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
  - [Body](/javascript/api/outlook/office.body) and all its child members
  - [Location](/javascript/api/outlook/office.location) and all its child members
  - [Recipients](/javascript/api/outlook/office.recipients) and all its child members
  - [Subject](/javascript/api/outlook/office.subject) and all its child members
  - [Time](/javascript/api/outlook/office.time) and all its child members

## read item permission

The **read item** permission is the next level of permission in the permissions model.

### Can do

- Read all the properties of the current item in a read or compose form. For example, [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) in a read form and [item.to.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) in a compose form.

- In Exchange on-premises environments, get a callback token to get the full mail item with Exchange Web Services (EWS).

    [!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

- [Write custom properties](/javascript/api/outlook/office.customproperties) set by the add-in on that item.

- Use regular expressions in a [contextual add-in](contextual-outlook-add-ins.md).

### Can't do

- Use the token provided by **mailbox.getCallbackTokenAsync** to:
  - Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.
  - Get the current calendar event item using the Outlook REST API.

- Use any of the following APIs.
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.bcc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.bcc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))
  - [item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))
  - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))
  - [item.cc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.cc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.end.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))
  - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.start.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))
  - [item.to.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.to.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))

## read/write item permission

Specify **read/write item** permission in the manifest to request this permission. Mail add-ins activated in compose forms that use write methods (**Message.to.addAsync** or **Message.to.setAsync**) must use at least this level of permission.

### Can do

- Read and write all item-level properties of the item that is being viewed or composed in Outlook.

- [Add or remove attachments](add-and-remove-attachments-to-an-item-in-a-compose-form.md) of that item.

- Use all other members of the Office JavaScript API that are applicable to mail add-ins, except **Mailbox.makeEWSRequestAsync**.

### Can't do

- Use the token provided by **mailbox.getCallbackTokenAsync** to:
  - Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.
  - Get the current calendar event item using the Outlook REST API.

- Use **mailbox.makeEWSRequestAsync**.

## read/write mailbox permission

The **read/write mailbox** permission is the highest level of permission.

In Exchange on-premises environments, the token provided by **mailbox.getCallbackTokenAsync** provides access to use Exchange Web Services (EWS) operations or Outlook REST APIs to do the following:

- Read and write all properties of any item in the user's mailbox.
- Create, read, and write to any folder or item in that mailbox.
- Send an item from that mailbox.

Through **mailbox.makeEwsRequestAsync**, you can access the following EWS operations.

- [CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)
- [CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)
- [CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)
- [FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)
- [FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)
- [FindItem](/exchange/client-developer/web-service-reference/finditem-operation)
- [GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)
- [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)
- [MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)
- [SendItem](/exchange/client-developer/web-service-reference/senditem-operation)
- [UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)

Attempting to use an unsupported operation will result in an error response.

[!INCLUDE [legacy-exchange-token-deprecation](../includes/legacy-exchange-token-deprecation.md)]

## See also

- [Privacy, permissions, and security for Outlook add-ins](../concepts/privacy-and-security.md)
