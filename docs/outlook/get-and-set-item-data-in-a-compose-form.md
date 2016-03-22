
# Get and set item data in a compose form in Outlook
Learn how to get or set various properties of an item in an Outlook add-in in a compose scenario, including its recipients, subject, body, and appointment location and time.




## Getting and setting item properties for a compose add-in


In a compose form, you can get most of the properties that are exposed on the same kind of item as in a read form (such as attendees, recipients, subject, and body), and you can get a few extra properties that are relevant in only a compose form but not a read form (body, bcc). 

For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The  [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) and [item.conversationId](../../reference/outlook/Office.context.mailbox.item.md) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.

Other than accessing item properties in the JavaScript API for Office, you can access item-level properties using Exchange Web Services (EWS). With the  **ReadWriteMailbox** permission, you can use the [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) method to access EWS operations, [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) and [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx), to get and set more properties of an item or items in the user's mailbox.  **makeEwsRequestAsync** is available in both compose and read forms. For more information about the **ReadWriteMailbox** permission, and accessing EWS through the Office Add-ins platform, see [Understanding Outlook add-in permissions](../outlook/understanding-outlook-add-in-permissions.md) and [Call web services from an Outlook add-in](../outlook/web-services.md).


**Table 1. Asynchronous methods to get or set item properties in a compose form**


|**Property**|**Property type**|**Asynchronous method to get**|**Asynchronous method(s) to set**|
|:-----|:-----|:-----|:-----|
|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|[Recipients](../../reference/outlook/Recipients.md)|[Recipients.getAsync](../../reference/outlook/Recipients.md)|[Recipients.addAsync](../../reference/outlook/Recipients.md)[Recipients.setAsync](../../reference/outlook/Recipients.md)|
|[body](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body.getAsync](../../reference/outlook/Body.md)|[Body.prependAsync](../../reference/outlook/Body.md)[Body.setAsync](../../reference/outlook/Body.md)[Body.setSelectedDataAsync](../../reference/outlook/Body.md)|
|[cc](../../reference/outlook/Office.context.mailbox.item.md)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../../reference/outlook/Office.context.mailbox.item.md)|[Time](../../reference/outlook/Time.md)|[Time.getAsync](../../reference/outlook/Time.md)|[Time.setAsync](../../reference/outlook/Time.md)|
|[location](../../reference/outlook/Office.context.mailbox.item.md)|[Location](../../reference/outlook/Location.md)|[Location.getAsync](../../reference/outlook/Location.md)|[Location.setAsync](../../reference/outlook/Location.md)|
|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../../reference/outlook/Office.context.mailbox.item.md)|Time|Time.getAsync|Time.setAsync|
|[subject](../../reference/outlook/Office.context.mailbox.item.md)|[Subject](../../reference/outlook/Subject.md)|[Subject.getAsync](../../reference/outlook/Subject.md)|[Subject.setAsync](../../reference/outlook/Subject.md)|
|[to](../../reference/outlook/Office.context.mailbox.item.md)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|



## Additional resources



- [Create Outlook add-ins for compose forms](../outlook/compose-scenario.md)
    
- [Understanding Outlook add-in permissions](../outlook/understanding-outlook-add-in-permissions.md)
    
- [Call web services from an Outlook add-in](../outlook/web-services.md)
    
- [Get and set Outlook item data in read or compose forms](../outlook/item-data.md)
    


