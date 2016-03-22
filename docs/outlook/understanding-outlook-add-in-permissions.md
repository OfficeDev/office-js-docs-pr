
# Understanding Outlook add-in permissions

Outlook add-ins specify the required permission level in their manifest. The available levels are  **Restricted**,  **ReadItem**,  **ReadWriteItem**, or  **ReadWriteMailbox**. These levels of permissions are cumulative:  **Restricted** is the lowest level, and each higher level includes the permissions of all the lower levels. **ReadWriteMailbox** includes all the supported permissions.

You can see the permissions requested by a mail add-in before installing it from the Office Store. You can also see the required permissions of installed add-ins in the Exchange Admin Center.


## Restricted permission


The  **Restricted** permission is the most basic level of permission. Specify **Restricted** in the [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) element in the manifest to request this permission. Outlook assigns this permission to a mail add-in by default if the add-in does not request a specific permission in its manifest.


### Can do


- [Get only specific entities](../outlook/match-strings-in-an-item-as-well-known-entities.md#MailAppEntities_Retrieving) (phone number, address, URL) from the item's subject or body.
    
- Specify an [ItemIs activation rule](../outlook/manifests/activation-rules.md#MailAppDefineRules_ItemIs) that requires the current item in a read or compose form to be a specific item type, or [ItemHasKnownEntity rule](../outlook/match-strings-in-an-item-as-well-known-entities.md#MailAppEntities_Activating) that matches any of a smaller subset of supported well-known entities (phone number, address, URL) in the selected item.
    
- Access any properties and methods that do  **not** pertain to specific information about the user or item. (See the next section for the list of members that do.)
    

### Can't do


- Use an [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) rule on the contact, email address, meeting suggestion, or task suggestion entitiy.
    
- Use the [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) or [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) rule.
    
- Access the members in the following list that pertain to the information of the user or item. Attempting to access members in this list will return  **null** and result in an error message which states that Outlook requires the mail add-in to have elevated permission.
    
      - [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.attachments](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.bcc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.body](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.cc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.from](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.organizer](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.resources](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.sender](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.to](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.userProfile](../../reference/outlook/Office.context.mailbox.userProfile.md)
    
  - [Body](../../reference/outlook/Body.md) and all its child members
    
  - [Location](../../reference/outlook/Location.md) and all its child members
    
  - [Recipients](../../reference/outlook/Recipients.md) and all its child members
    
  - [Subject](../../reference/outlook/Subject.md) and all its child members
    
  - [Time](../../reference/outlook/Time.md) and all its child members
    

## ReadItem permission


The  **ReadItem** permission is the next level of permission in the permissions model. Specify **ReadItem** in the **Permissions** element in the manifest to request this permission.


### Can do


- [Read all the properties](../outlook/item-data.md) of the current item in a read or [compose form](../outlook/get-and-set-item-data-in-a-compose-form.md), for example, [item.to](../../reference/outlook/Office.context.mailbox.item.md) in a read form and [item.to.getAsync](../../reference/outlook/Recipients.md) in a compose form.
    
- [Get a callback token to get item attachments](../outlook/get-attachments-of-an-outlook-item.md) or the full item.
    
- [Write custom properties](http://msdn.microsoft.com/library/30217d63-7615-4f3f-8618-c91e4e60cd43%28Office.15%29.aspx) set by the add-in on that item.
    
- [Get all existing well-known entities](../outlook/match-strings-in-an-item-as-well-known-entities.md#MailAppEntities_Retrieving), not just a subset, from the item's subject or body.
    
- Use all the [well-known entities](../outlook/manifests/activation-rules.md#MailAppDefineRules_ItemHasKnownEntity) in [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) rules, or [regular expressions](../outlook/manifests/activation-rules.md#MailAppDefineRules_ItemHasRegularExpressionMatch) in [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) rules. The following example follows schema v1.1. It shows a rule that activates the add-in if one or more of the well-known entities are found in the subject or body of the selected message:
    

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

Access  **mailbox.makeEWSRequestAsync** or the following write methods:


- [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.bcc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.bcc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.body.prependAsync](../../reference/outlook/Body.md)
    
- [item.body.setAsync](../../reference/outlook/Body.md)
    
- [item.body.setSelectedDataAsync](../../reference/outlook/Body.md)
    
- [item.cc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.cc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.end.setAsync](../../reference/outlook/Time.md)
    
- [item.location.setAsync](../../reference/outlook/Location.md)
    
- [item.optionalAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.optionalAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.requiredAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.requiredAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.start.setAsync](../../reference/outlook/Time.md)
    
- [item.subject.setAsync](../../reference/outlook/Subject.md)
    
- [item.to.addAsync](../../reference/outlook/Recipients.md)
    
- [item.to.setAsync](../../reference/outlook/Recipients.md)
    

## ReadWriteItem permission


Specify  **ReadWriteItem** in the **Permissions** element in the manifest to request this permission. Mail add-ins activated in compose forms that use write methods ( **Message.to.addAsync** or **Message.to.setAsync**) must use at least this level of permission.


### Can do


- [Read and write all item-level properties](../outlook/item-data.md) of the item that is being viewed or composed in Outlook.
    
- [Add or remove attachments](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md) of that item.
    
- Use all other members of the JavaScript API for Office that are applicable to mail add-ins, except  **Mailbox.makeEWSRequestAsync**.
    

### Can't do

Use  **Mailbox.makeEWSRequestAsync**.


## ReadWriteMailbox permission


The  **ReadWriteMailbox** permission is the highest level of permission. Specify **ReadWriteMailbox** in the **Permissions** element in the manifest to request this permission.

In addition to what the  **ReadWriteItem** permission supports, by using **Mailbox.makeEWSRequestAsync**, you can access supported Exchange Web Services (EWS) operations to do the following:


- Read and write all properties of any item in the user's mailbox.
    
- Create, read, and write to any folder or item in that mailbox.
    
- Send an item from that mailbox
    
Through  **mailbox.makeEWSRequestAsync**, you can access the following EWS operations:


- [CopyItem](http://msdn.microsoft.com/en-us/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)
    
- [CreateFolder](http://msdn.microsoft.com/en-us/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)
    
- [CreateItem](http://msdn.microsoft.com/en-us/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)
    
- [FindConversation](http://msdn.microsoft.com/en-us/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)
    
- [FindFolder](http://msdn.microsoft.com/en-us/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)
    
- [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)
    
- [GetConversationItems](http://msdn.microsoft.com/en-us/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)
    
- [GetFolder](http://msdn.microsoft.com/en-us/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)
    
- [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)
    
- [MarkAsJunk](http://msdn.microsoft.com/en-us/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)
    
- [MoveItem](http://msdn.microsoft.com/en-us/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)
    
- [SendItem](http://msdn.microsoft.com/en-us/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)
    
- [UpdateFolder](http://msdn.microsoft.com/en-us/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)
    
- [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)
    
Attempting to use an unsupported operation will result in an error response.


## Additional resources



- [Privacy, permissions, and security for Outlook add-ins](../outlook/../../docs/develop/privacy-and-security.md)
    
- [Match strings in an Outlook item as well-known entities](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
