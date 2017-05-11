# Outlook add-in API requirement set 1.1

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a [requirement set](../tutorial-api-requirement-sets.md) other than the latest requirement set. 

## What's new in 1.1?

Requirement set 1.1 includes all of the features of Requirement set 1.0. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.

### Change log

- Added [Body](https://dev.office.com/reference/add-ins/outlook/1.1/Body?product=outlook&version=v1.1) object: Provides methods for adding and updating the content of an item in an Outlook add-in.
- Added [Location](https://dev.office.com/reference/add-ins/outlook/1.1/Location?product=outlook&version=v1.1) object: Provides methods to get and set the location of a meeting in an Outlook add-in.
- Added [Recipients](https://dev.office.com/reference/add-ins/outlook/1.1/Recipients?product=outlook&version=v1.1) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.
- Added [Subject](https://dev.office.com/reference/add-ins/outlook/1.1/Subject?product=outlook&version=v1.1) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.
- Added [Time](https://dev.office.com/reference/add-ins/outlook/1.1/Time?product=outlook&version=v1.1) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.
- Added [Office.context.mailbox.item.addFileAttachmentAsync](https://dev.office.com/reference/add-ins/outlook/1.1/Office.context.mailbox.item?product=outlook&version=v1.1#addfileattachmentasyncuri-attachmentname-options-callback): Adds a file to a message or appointment as an attachment.
- Added [Office.context.mailbox.item.addItemAttachmentAsync](https://dev.office.com/reference/add-ins/outlook/1.1/Office.context.mailbox.item?product=outlook&version=v1.1#additemattachmentasyncitemid-attachmentname-options-callback): Adds an Exchange item, such as a message, as an attachment to the message or appointment.
- Added [Office.context.mailbox.item.removeAttachmentAsync](https://dev.office.com/reference/add-ins/outlook/1.1/Office.context.mailbox.item?product=outlook&version=v1.1#removeattachmentasyncattachmentid-options-callback): Removes an attachment from a message or appointment.
- Added [Office.context.mailbox.item.body](https://dev.office.com/reference/add-ins/outlook/1.1/Office.context.mailbox.item?product=outlook&version=v1.1#body-body): Gets an object that provides methods for manipulating the body of an item.
- Added [Office.context.mailbox.item.bcc](https://dev.office.com/reference/add-ins/outlook/1.1/Office.context.mailbox.item?product=outlook&version=v1.1#bcc-recipients): Gets or sets the recipients on the Bcc (blind carbon copy) line of a message.
- Added [Office.MailboxEnums.RecipientType](https://dev.office.com/reference/add-ins/outlook/1.1/Office.MailboxEnums?product=outlook&version=v1.1#recipienttype-string): Specifies the type of recipient for an appointment.

## Additional resources

- [Outlook add-ins](../../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://developer.microsoft.com/en-us/outlook/code-samples)
- [Get started](https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial)
