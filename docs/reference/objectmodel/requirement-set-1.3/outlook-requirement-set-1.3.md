# Outlook add-in API requirement set 1.3

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](/javascript/office/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set. 

## What's new in 1.3?

Requirement set 1.3 includes all of the features of [Requirement set 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). It added the following features.

- Added support for [add-in commands](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook).
- Added ability to save or close an item being composed.
- Enhanced [Body](/javascript/api/outlook_1_3/office.body) object to allow addins to get or set the entire body.
- Added conversion methods to convert IDs between EWS and REST formats.
- Added ability to add notification messages to the info bar on items.

### Change log

- Added [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-): Returns the current body in a specified format.
- Added [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-): Replaces the entire body with the specified text.
- Added [Office.context.officeTheme](office.context.md#officetheme-object): Provides access to the Office theme colors.
- Added [Event](/javascript/api/office/office.addincommands.event) object: Passed as a parameter to UI-less command functions in an Outlook add-in. Used to signal completion of processing.
- Added [Office.context.mailbox.item.close](office.context.mailbox.item.md#close): Closes the current item that is being composed.
- Added [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback): Asynchronously saves an item.
- Added [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages): Gets the notification messages for an item.
- Added [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string): Converts an item ID formatted for REST into EWS format.
- Added [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string): Converts an item ID formatted for EWS into REST format.
- Added [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype): Specifies the notification message type for an appointment or message.
- Added [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion): Specifies the version of the REST API that corresponds to a REST-formatted item ID.
- Added [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) object: Provides methods for accessing notification messages in an Outlook add-in.
- Added [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) type: Returned by the `NotificationMessages.getAllAsync` method.

## See also

- [Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](https://docs.microsoft.com/outlook/add-ins/quick-start)