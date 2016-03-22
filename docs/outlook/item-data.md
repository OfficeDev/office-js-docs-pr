
# Get and set Outlook item data in read or compose forms

Starting in version 1.1 of the Office Add-ins manifests schema, Outlook can activate add-ins when the user is viewing or composing an item. Depending on whether an add-in is activated in a read or compose form, the properties that are available to the add-in on the item differ as well. For example, the [dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) and [dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md) properties are defined only for an item that has already been sent (item is subsequently viewed in a read form) but not when the item is being created (in a compose form). Another example is the [bcc](../../reference/outlook/Office.context.mailbox.item.md) property which is only meaningful when a message is being authored (in a compose form), and is not accessible to the user in a read form.

Table 1 shows the item-level properties in the JavaScript API for Office that are available in each of read and compose modes of mail add-ins. Typically, those properties available in read forms are read-only, and those available in compose forms are read/write, with the exception of the [itemId](../../reference/outlook/Office.context.mailbox.item.md) and [conversationId](../../reference/outlook/Office.context.mailbox.item.md) properties, which are always read-only regardless. For the remaining item-level properties available in compose forms, because the add-in and user can possibly be reading or writing the same property at the same time, the methods to get or set them in compose mode are asynchronous, and hence the type of the objects returned by these properties are also different in compose forms than in read forms. For more information about using asynchronous methods to get or set item-level properties in compose mode, see [Get and set item data in a compose form in Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md).


**Table 1. Item properties available in compose and read forms**


|**Item type**|**Property**|**Property type in read forms**|**Property type in compose forms**|
|:-----|:-----|:-----|:-----|
|Appointments and messages|[dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript  **Date** object|Property not available|
|Appointments and messages|[dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript  **Date** object|Property not available|
|Appointments and messages|[itemClass](../../reference/outlook/Office.context.mailbox.item.md)|String|Property not available|
|Appointments and messages|[itemId](../../reference/outlook/Office.context.mailbox.item.md)|String|Property not available|
|Appointments and messages|[itemType](../../reference/outlook/Office.context.mailbox.item.md)|String in [ItemType](../../reference/outlook/Office.MailboxEnums.md) enumeration|Property not available|
|Appointments and messages|[attachments](../../reference/outlook/Office.context.mailbox.item.md)|[AttachmentDetails](../../reference/outlook/simple-types.md)|Property not available|
|Appointments and messages|[body](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body](../../reference/outlook/Body.md)|
|Appointments|[end](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript  **Date** object|[Time](../../reference/outlook/Time.md)|
|Appointments|[location](../../reference/outlook/Office.context.mailbox.item.md)|String|[Location](../../reference/outlook/Location.md)|
|Appointments and messages|[normalizedSubject](../../reference/outlook/Office.context.mailbox.item.md)|String|Property not available|
|Appointments|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|[EmailAddressDetails](../../reference/outlook/simple-types.md)|[Recipients](../../reference/outlook/Recipients.md)|
|Appointments|[organizer](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Property not available|
|Appointments|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Recipients|
|Appointments|[resources](../../reference/outlook/Office.context.mailbox.item.md)|String|Property not available|
|Appointments|[start](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript  **Date** object|Time|
|Appointments and messages|[subject](../../reference/outlook/Office.context.mailbox.item.md)|String|[Subject](../../reference/outlook/Subject.md)|
|Messages|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|Property not available|Recipients|
|Messages|[cc](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Recipients|
|Messages|[conversationId](../../reference/outlook/Office.context.mailbox.item.md)|String|String (read only)|
|Messages|[from](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Property not available|
|Messages|[internetMessageId](../../reference/outlook/Office.context.mailbox.item.md)|Integer|Property not available|
|Messages|[sender](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Property not available|
|Messages|[to](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Recipients|

## Using Exchange Server callback tokens from a read add-in


If your Outlook add-in is activated in read forms, you can get an Exchange callback token. This token can be usedin server-side code to access the full item via Exchange Web Services (EWS). By specifying the  **ReadItem** permission in the add-in manifest, you can use the [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) method to get an Exchange callback token, the [mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) property to get the URL of the EWS endpoint for the user's mailbox, and [item.itemId](../../reference/outlook/Office.context.mailbox.item.md) to get the EWS ID for the selected item. You can then pass the callback token, EWS endpoint URL, and the EWS item ID to server-side code to access the [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) operation, to get more properties of the item.


## Accessing EWS from a read or compose add-in


You can also use the [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) method to access the Exchange Web Services (EWS) operations [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) and [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) directly from the add-in. You can use these operations to get and set many properties of a specified item. This method is available to Outlook add-ins regardless of whether the add-in has been activated in a read or compose form, as long as you specify the **ReadWriteMailbox** permission in the add-in manifest. For more information in using **makeEwsRequestAsync** to access EWS operations, see [Call web services from an Outlook add-in](../outlook/web-services.md).


## Additional resources



- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
- [Get and set item data in a compose form in Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Call web services from an Outlook add-in](../outlook/web-services.md)
    


