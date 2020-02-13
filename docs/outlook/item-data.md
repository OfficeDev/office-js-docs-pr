---
title: Get or set item data in an Outlook add-in
description: Depending on whether an add-in is activated in a read or compose form, the properties that are available to the add-in on an item differ.
ms.date: 12/10/2019
localization_priority: Normal
---

# Get and set Outlook item data in read or compose forms

Starting in version 1.1 of the Office Add-ins manifests schema, Outlook can activate add-ins when the user is viewing or composing an item. Depending on whether an add-in is activated in a read or compose form, the properties that are available to the add-in on the item differ as well.

For example, the [dateTimeCreated](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties) and [dateTimeModified](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties) properties are defined only for an item that has already been sent (item is subsequently viewed in a read form) but not when the item is being created (in a compose form). Another example is the [bcc](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties) property, which is only meaningful when a message is being authored (in a compose form), and is not accessible to the user in a read form.

## Item properties available in compose and read forms

Table 1 shows the item-level properties in the JavaScript API for Office that are available in each mode (read and compose) of mail add-ins. Typically, those properties available in read forms are read-only, and those available in compose forms are read/write, with the exception of the [itemId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties), [conversationId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties), and [itemType](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties) properties, which are always read-only regardless.

For the remaining item-level properties available in compose forms, because the add-in and user can possibly be reading or writing the same property at the same time, the methods to get or set them in compose mode are asynchronous, and hence the type of the objects returned by these properties may also be different in compose forms than in read forms. For more information about using asynchronous methods to get or set item-level properties in compose mode, see [Get and set item data in a compose form in Outlook](get-and-set-item-data-in-a-compose-form.md).


**Table 1. Item properties available in compose and read forms**

<br/>

|**Item type**|**Property**|**Property type in read forms**|**Property type in compose forms**|
|:-----|:-----|:-----|:-----|
|Appointments and messages|[dateTimeCreated](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** object|Property not available|
|Appointments and messages|[dateTimeModified](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** object|Property not available|
|Appointments and messages|[itemClass](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|String|Property not available|
|Appointments and messages|[itemId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|String|Property not available|
|Appointments and messages|[itemType](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|String in [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) enumeration|String in [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) enumeration (read only)|
|Appointments and messages|[attachments](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|Property not available|
|Appointments and messages|[body](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[Body](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|Appointments and messages|[normalizedSubject](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|String|Property not available|
|Appointments and messages|[subject](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|String|[Subject](/javascript/api/outlook/office.subject)|
|Appointments|[end](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** object|[Time](/javascript/api/outlook/office.time)|
|Appointments|[location](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|String|[Location](/javascript/api/outlook/office.location)|
|Appointments|[optionalAttendees](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Appointments|[organizer](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)|
|Appointments|[requiredAttendees](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Appointments|[start](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** object|[Time](/javascript/api/outlook/office.time)|
|Messages|[bcc](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|Property not available|[Recipients](/javascript/api/outlook/office.recipients)|
|Messages|[cc](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|Messages|[conversationId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|String|String (read only)|
|Messages|[from](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)|
|Messages|[internetMessageId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|Integer|Property not available|
|Messages|[sender](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|Property not available|
|Messages|[to](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|

## Use Exchange Server callback tokens from a read add-in

If your Outlook add-in is activated in read forms, you can get an Exchange callback token. This token can be used in server-side code to access the full item via Exchange Web Services (EWS).

By specifying the **ReadItem** permission in the add-in manifest, you can use the [mailbox.getCallbackTokenAsync](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods) method to get an Exchange callback token, the [mailbox.ewsUrl](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#properties) property to get the URL of the EWS endpoint for the user's mailbox, and [item.itemId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties) to get the EWS ID for the selected item. You can then pass the callback token, EWS endpoint URL, and the EWS item ID to server-side code to access the [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation, to get more properties of the item.


## Access EWS from a read or compose add-in

You can also use the [mailbox.makeEwsRequestAsync](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods) method to access the Exchange Web Services (EWS) operations [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) and [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) directly from the add-in. You can use these operations to get and set many properties of a specified item. This method is available to Outlook add-ins regardless of whether the add-in has been activated in a read or compose form, as long as you specify the **ReadWriteMailbox** permission in the add-in manifest.

For more information about using **makeEwsRequestAsync** to access EWS operations, see [Call web services from an Outlook add-in](web-services.md).


## See also

- [Get and set item data in a compose form in Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Call web services from an Outlook add-in](web-services.md)
