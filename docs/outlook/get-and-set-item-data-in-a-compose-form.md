---
title: Get and set item data in a compose form in Outlook | Microsoft Docs
description: Get or set various properties of an item in an Outlook add-in in a compose scenario, including its recipients, subject, body, and appointment location and time.
author: jasonjoh
ms.topic: article
ms.technology: office-add-ins
ms.date: 08/09/2017
ms.author: jasonjoh
---

# Get and set item data in a compose form in Outlook

Learn how to get or set various properties of an item in an Outlook add-in in a compose scenario, including its recipients, subject, body, and appointment location and time.

## Getting and setting item properties for a compose add-in

In a compose form, you can get most of the properties that are exposed on the same kind of item as in a read form (such as attendees, recipients, subject, and body), and you can get a few extra properties that are relevant in only a compose form but not a read form (body, bcc).

For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The  [item.itemType](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) and [item.conversationId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#nullable-conversationid-string) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.

Other than accessing item properties in the JavaScript API for Office, you can access item-level properties using Exchange Web Services (EWS). With the **ReadWriteMailbox** permission, you can use the [mailbox.makeEwsRequestAsync](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#makeewsrequestasyncdata-callback-usercontext) method to access EWS operations, [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) and [UpdateItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/updateitem-operation), to get and set more properties of an item or items in the user's mailbox.

The `makeEwsRequestAsync` function is available in both compose and read forms. For more information about the **ReadWriteMailbox** permission, and accessing EWS through the Office Add-ins platform, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md) and [Call web services from an Outlook add-in](web-services.md).

**Table 1. Asynchronous methods to get or set item properties in a compose form**

<br/>

| Property | Property type | Asynchronous method to get | Asynchronous method(s) to set |
|:-----|:-----|:-----|:-----|
|[bcc](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#bcc-recipientsjavascriptapioutlook15officerecipients)|[Recipients](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Recipients)|[Recipients.getAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Recipients#getasync-options--callback-)|[Recipients.addAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Recipients#addasync-recipients--options--callback-), [Recipients.setAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Recipients#setasync-recipients--options--callback-)|
|[body](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#body-bodyjavascriptapioutlook15officebody)|[Body](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Body)|[Body.getAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Body#getasync-coerciontype--options--callback-)|[Body.prependAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Body#prependasync-data--options--callback-), [Body.setAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Body#setasync-data--options--callback-), [Body.setSelectedDataAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Body#setselecteddataasync-data--options--callback-)|
|[cc](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#end-datetimejavascriptapioutlook15officetime)|[Time](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Time)|[Time.getAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Time#getasync-options--callback-)|[Time.setAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Time#setasync-datetime--options--callback-)|
|[location](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#location-stringlocationjavascriptapioutlook15officelocation)|[Location](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Location)|[Location.getAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Location#getasync-options--callback-)|[Location.setAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Location#setasync-location--options--callback-)|
|[optionalAttendees](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#start-datetimejavascriptapioutlook15officetime)|Time|Time.getAsync|Time.setAsync|
|[subject](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#subject-stringsubjectjavascriptapioutlook15officesubject)|[Subject](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Subject)|[Subject.getAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Subject#getasync-options--callback-)|[Subject.setAsync](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Subject#setasync-subject--options--callback-)|
|[to](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## See also

- [Create Outlook add-ins for compose forms](compose-scenario.md)
- [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md)
- [Call web services from an Outlook add-in](web-services.md)
- [Get and set Outlook item data in read or compose forms](item-data.md)