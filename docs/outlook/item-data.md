---
title: Get or set item data in an Outlook add-in | Microsoft Docs
description: Depending on whether an add-in is activated in a read or compose form, the properties that are available to the add-in on an item differ.
author: jasonjoh
ms.topic: article
ms.technology: office-add-ins
ms.date: 06/13/2017
ms.author: jasonjoh
---

# Get and set Outlook item data in read or compose forms

Starting in version 1.1 of the Office Add-ins manifests schema, Outlook can activate add-ins when the user is viewing or composing an item. Depending on whether an add-in is activated in a read or compose form, the properties that are available to the add-in on the item differ as well.

For example, the [dateTimeCreated](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#datetimecreated-date) and [dateTimeModified](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#datetimemodified-date) properties are defined only for an item that has already been sent (item is subsequently viewed in a read form) but not when the item is being created (in a compose form). Another example is the [bcc](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#bcc-recipientsjavascriptapioutlook15officerecipients) property, which is only meaningful when a message is being authored (in a compose form), and is not accessible to the user in a read form.

## Item properties available in compose and read forms

Table 1 shows the item-level properties in the JavaScript API for Office that are available in each of read and compose modes of mail add-ins. Typically, those properties available in read forms are read-only, and those available in compose forms are read/write, with the exception of the [itemId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#nullable-itemid-string) and [conversationId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#nullable-conversationid-string) properties, which are always read-only regardless.

For the remaining item-level properties available in compose forms, because the add-in and user can possibly be reading or writing the same property at the same time, the methods to get or set them in compose mode are asynchronous, and hence the type of the objects returned by these properties are also different in compose forms than in read forms. For more information about using asynchronous methods to get or set item-level properties in compose mode, see [Get and set item data in a compose form in Outlook](get-and-set-item-data-in-a-compose-form.md).


**Table 1. Item properties available in compose and read forms**

<br/>

|**Item type**|**Property**|**Property type in read forms**|**Property type in compose forms**|
|:-----|:-----|:-----|:-----|
|Appointments and messages|[dateTimeCreated](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#datetimecreated-date)|JavaScript **Date** object|Property not available|
|Appointments and messages|[dateTimeModified](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#datetimemodified-date)|JavaScript **Date** object|Property not available|
|Appointments and messages|[itemClass](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#itemclass-string)|String|Property not available|
|Appointments and messages|[itemId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#nullable-itemid-string)|String|Property not available|
|Appointments and messages|[itemType](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype)|String in [ItemType](https://docs.microsoft.com/javascript/api/outlook_1_5/office.mailboxenums.itemtype) enumeration|Property not available|
|Appointments and messages|[attachments](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails)|[AttachmentDetails](https://docs.microsoft.com/javascript/api/outlook_1_5/office.attachmentdetails)|Property not available|
|Appointments and messages|[body](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#body-bodyjavascriptapioutlook15officebody)|[Body](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Body)|[Body](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Body)|
|Appointments|[end](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#end-datetimejavascriptapioutlook15officetime)|JavaScript **Date** object|[Time](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Time)|
|Appointments|[location](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#location-stringlocationjavascriptapioutlook15officelocation)|String|[Location](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Location)|
|Appointments and messages|[normalizedSubject](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#normalizedsubject-string)|String|Property not available|
|Appointments|[optionalAttendees](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients)|[EmailAddressDetails](https://docs.microsoft.com/javascript/api/outlook_1_5/office.emailaddressdetails)|[Recipients](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Recipients)|
|Appointments|[organizer](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails)|EmailAddressDetails|Property not available|
|Appointments|[requiredAttendees](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients)|EmailAddressDetails|Recipients|
|Appointments|[start](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#start-datetimejavascriptapioutlook15officetime)|JavaScript **Date** object|Time|
|Appointments and messages|[subject](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#subject-stringsubjectjavascriptapioutlook15officesubject)|String|[Subject](https://docs.microsoft.com/javascript/api/outlook_1_5/office.Subject)|
|Messages|[bcc](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#bcc-recipientsjavascriptapioutlook15officerecipients)|Property not available|Recipients|
|Messages|[cc](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients)|EmailAddressDetails|Recipients|
|Messages|[conversationId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#nullable-conversationid-string)|String|String (read only)|
|Messages|[from](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails)|EmailAddressDetails|Property not available|
|Messages|[internetMessageId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#internetmessageid-string)|Integer|Property not available|
|Messages|[sender](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails)|EmailAddressDetails|Property not available|
|Messages|[to](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients)|EmailAddressDetails|Recipients|

## Use Exchange Server callback tokens from a read add-in

If your Outlook add-in is activated in read forms, you can get an Exchange callback token. This token can be used in server-side code to access the full item via Exchange Web Services (EWS).

By specifying the **ReadItem** permission in the add-in manifest, you can use the [mailbox.getCallbackTokenAsync](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#getcallbacktokenasyncoptions-callback) method to get an Exchange callback token, the [mailbox.ewsUrl](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#ewsurl-string) property to get the URL of the EWS endpoint for the user's mailbox, and [item.itemId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#nullable-itemid-string) to get the EWS ID for the selected item. You can then pass the callback token, EWS endpoint URL, and the EWS item ID to server-side code to access the [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation, to get more properties of the item.


## Access EWS from a read or compose add-in

You can also use the [mailbox.makeEwsRequestAsync](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#makeewsrequestasyncdata-callback-usercontext) method to access the Exchange Web Services (EWS) operations [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) and [UpdateItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/updateitem-operation) directly from the add-in. You can use these operations to get and set many properties of a specified item. This method is available to Outlook add-ins regardless of whether the add-in has been activated in a read or compose form, as long as you specify the **ReadWriteMailbox** permission in the add-in manifest.

For more information about using **makeEwsRequestAsync** to access EWS operations, see [Call web services from an Outlook add-in](web-services.md).


## See also

- [Get and set item data in a compose form in Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Call web services from an Outlook add-in](web-services.md)
    


