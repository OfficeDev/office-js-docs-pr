---
title: Get and set item data in a compose form in Outlook
description: Get or set various properties of an item in an Outlook add-in in a compose scenario, including its recipients, subject, body, and appointment location and time.
ms.date: 12/10/2019
localization_priority: Normal
---

# Get and set item data in a compose form in Outlook

Learn how to get or set various properties of an item in an Outlook add-in in a compose scenario, including its recipients, subject, body, and appointment location and time.

## Getting and setting item properties for a compose add-in

In a compose form, you can get most of the properties that are exposed on the same kind of item as in a read form (such as attendees, recipients, subject, and body), and you can get a few extra properties that are relevant in only a compose form but not a read form (body, bcc).

For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The  [item.itemType](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties) and [item.conversationId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.

Other than accessing item properties in the JavaScript API for Office, you can access item-level properties using Exchange Web Services (EWS). With the **ReadWriteMailbox** permission, you can use the [mailbox.makeEwsRequestAsync](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods) method to access EWS operations, [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) and [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation), to get and set more properties of an item or items in the user's mailbox.

The `makeEwsRequestAsync` function is available in both compose and read forms. For more information about the **ReadWriteMailbox** permission, and accessing EWS through the Office Add-ins platform, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md) and [Call web services from an Outlook add-in](web-services.md).

**Table 1. Asynchronous methods to get or set item properties in a compose form**

<br/>

| Property | Property type | Asynchronous method to get | Asynchronous method(s) to set |
|:-----|:-----|:-----|:-----|
|[bcc](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[Recipients](/javascript/api/outlook/office.Recipients)|[Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)|[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)|
|[body](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[Body](/javascript/api/outlook/office.Body)|[Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)|[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-), [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)|
|[cc](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[Time](/javascript/api/outlook/office.Time)|[Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-)|[Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)|
|[location](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[Location](/javascript/api/outlook/office.Location)|[Location.getAsync](/javascript/api/outlook/office.Location#getasync-options--callback-)|[Location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-)|
|[optionalAttendees](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|Time|Time.getAsync|Time.setAsync|
|[subject](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|[Subject](/javascript/api/outlook/office.Subject)|[Subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-)|[Subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)|
|[to](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## See also

- [Create Outlook add-ins for compose forms](compose-scenario.md)
- [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md)
- [Call web services from an Outlook add-in](web-services.md)
- [Get and set Outlook item data in read or compose forms](item-data.md)
