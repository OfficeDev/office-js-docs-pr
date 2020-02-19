---
title: Understanding Outlook add-in permissions
description: Outlook add-ins specify the required permission level in their manifest, which include Restricted, ReadItem, ReadWriteItem, or ReadWriteMailbox. 
ms.date: 12/10/2019
localization_priority: Normal
---

# Understanding Outlook add-in permissions

Outlook add-ins specify the required permission level in their manifest. The available levels are **Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**. These levels of permissions are cumulative: **Restricted** is the lowest level, and each higher level includes the permissions of all the lower levels. **ReadWriteMailbox** includes all the supported permissions.

You can see the permissions requested by a mail add-in before installing it from [AppSource](https://appsource.microsoft.com). You can also see the required permissions of installed add-ins in the Exchange Admin Center.

## Restricted permission

The **Restricted** permission is the most basic level of permission. Specify **Restricted** in the [Permissions](../reference/manifest/permissions.md) element in the manifest to request this permission. Outlook assigns this permission to a mail add-in by default if the add-in does not request a specific permission in its manifest.

### Can do

- [Get only specific entities](match-strings-in-an-item-as-well-known-entities.md) (phone number, address, URL) from the item's subject or body.

- Specify an [ItemIs activation rule](activation-rules.md#itemis-rule) that requires the current item in a read or compose form to be a specific item type, or [ItemHasKnownEntity rule](match-strings-in-an-item-as-well-known-entities.md) that matches any of a smaller subset of supported well-known entities (phone number, address, URL) in the selected item.

- Access any properties and methods that do **not** pertain to specific information about the user or item (see the next section for the list of members that do).

### Can't do

- Use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule on the contact, email address, meeting suggestion, or task suggestion entitiy.

- Use the [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) or [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule.

- Access the members in the following list that pertain to the information of the user or item. Attempting to access members in this list will return **null** and result in an error message which states that Outlook requires the mail add-in to have elevated permission.

    - [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.from](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.organizer](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.sender](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.userProfile](../reference/objectmodel/preview-requirement-set/office.context.mailbox.userprofile.md)
    - [Body](/javascript/api/outlook/office.body) and all its child members
    - [Location](/javascript/api/outlook/office.location) and all its child members
    - [Recipients](/javascript/api/outlook/office.recipients) and all its child members
    - [Subject](/javascript/api/outlook/office.subject) and all its child members
    - [Time](/javascript/api/outlook/office.time) and all its child members

## ReadItem permission

The **ReadItem** permission is the next level of permission in the permissions model. Specify **ReadItem** in the **Permissions** element in the manifest to request this permission.

### Can do

- [Read all the properties](item-data.md) of the current item in a read or [compose form](get-and-set-item-data-in-a-compose-form.md), for example, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) in a read form and [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) in a compose form.

- [Get a callback token to get item attachments](get-attachments-of-an-outlook-item.md) or the full item with Exchange Web Services (EWS) or [Outlook REST APIs](use-rest-api.md).

- [Write custom properties](/javascript/api/outlook/office.CustomProperties) set by the add-in on that item.

- [Get all existing well-known entities](match-strings-in-an-item-as-well-known-entities.md), not just a subset, from the item's subject or body.

- Use all the [well-known entities](activation-rules.md#itemhasknownentity-rule) in [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rules, or [regular expressions](activation-rules.md#itemhasregularexpressionmatch-rule) in [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules. The following example follows schema v1.1. It shows a rule that activates the add-in if one or more of the well-known entities are found in the subject or body of the selected message:

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### Can't do

- Use the token provided by **mailbox.getCallbackTokenAsync** to:
    - Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.
    - Get the current calendar event item using the Outlook REST API.

- Use any of the following APIs:
    - [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.bcc.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.bcc.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [item.body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [item.cc.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.cc.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.end.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [item.location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.start.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [item.subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [item.to.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.to.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## ReadWriteItem permission

Specify **ReadWriteItem** in the **Permissions** element in the manifest to request this permission. Mail add-ins activated in compose forms that use write methods (**Message.to.addAsync** or **Message.to.setAsync**) must use at least this level of permission.

### Can do

- [Read and write all item-level properties](item-data.md) of the item that is being viewed or composed in Outlook.

- [Add or remove attachments](add-and-remove-attachments-to-an-item-in-a-compose-form.md) of that item.

- Use all other members of the JavaScript API for Office that are applicable to mail add-ins, except **Mailbox.makeEWSRequestAsync**.

### Can't do

- Use the token provided by **mailbox.getCallbackTokenAsync** to:
    - Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.
    - Get the current calendar event item using the Outlook REST API.

- Use **mailbox.makeEWSRequestAsync**.

## ReadWriteMailbox permission

The **ReadWriteMailbox** permission is the highest level of permission. Specify **ReadWriteMailbox** in the **Permissions** element in the manifest to request this permission.

In addition to what the **ReadWriteItem** permission supports, the token provided by **mailbox.getCallbackTokenAsync** provides access to use Exchange Web Services (EWS) operations or Outlook REST APIs to do the following:

- Read and write all properties of any item in the user's mailbox.
- Create, read, and write to any folder or item in that mailbox.
- Send an item from that mailbox

Through **mailbox.makeEWSRequestAsync**, you can access the following EWS operations:

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

## See also

- [Privacy, permissions, and security for Outlook add-ins](../develop/privacy-and-security.md)
- [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md)
