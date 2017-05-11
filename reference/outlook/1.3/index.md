# Outlook add-in API requirement set 1.3

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a [requirement set](../tutorial-api-requirement-sets.md) other than the latest requirement set. 

## What's new in 1.3?

Requirement set 1.3 includes all of the features of [Requirement set 1.2](../1.2/index.md). It added the following features.

- Added support for [add-in commands](../../../docs/outlook/add-in-commands-for-outlook.md).
- Added ability to save or close an item being composed.
- Enhanced [Body](https://dev.office.com/reference/add-ins/outlook/1.3/Body?product=outlook&version=v1.3) object to allow addins to get or set the entire body.
- Added conversion methods to convert IDs between EWS and REST formats.
- Added ability to add notification messages to the info bar on items.

### Change log

- Added [Body.getAsync](https://dev.office.com/reference/add-ins/outlook/1.3/Body?product=outlook&version=v1.3#getasynccoerciontype-options-callback): Returns the current body in a specified format.
- Added [Body.setAsync](https://dev.office.com/reference/add-ins/outlook/1.3/Body?product=outlook&version=v1.3#setasyncdata-options-callback): Replaces the entire body with the specified text.
- Added [Office.context.officeTheme](https://dev.office.com/reference/add-ins/outlook/1.3/Office.context?product=outlook&version=v1.3#officetheme-object): Provides access to the Office theme colors.
- Added [Event](https://dev.office.com/reference/add-ins/outlook/1.3/Event?product=outlook&version=v1.3) object: Passed as a parameter to UI-less command functions in an Outlook add-in. Used to signal completion of processing.
- Added [Office.context.mailbox.item.close](https://dev.office.com/reference/add-ins/outlook/1.3/Office.context.mailbox.item?product=outlook&version=v1.3#close): Closes the current item that is being composed.
- Added [Office.context.mailbox.item.saveAsync](https://dev.office.com/reference/add-ins/outlook/1.3/Office.context.mailbox.item?product=outlook&version=v1.3#saveasyncoptions-callback): Asynchronously saves an item.
- Added [Office.context.mailbox.item.notificationMessages](https://dev.office.com/reference/add-ins/outlook/1.3/Office.context.mailbox.item?product=outlook&version=v1.3#notificationmessages-notificationmessages): Gets the notification messages for an item.
- Added [Office.context.mailbox.convertToEwsId](https://dev.office.com/reference/add-ins/outlook/1.3/Office.context.mailbox?product=outlook&version=v1.3#converttoewsiditemid-restversion--string): Converts an item ID formatted for REST into EWS format.
- Added [Office.context.mailbox.convertToRestId](https://dev.office.com/reference/add-ins/outlook/1.3/Office.context.mailbox?product=outlook&version=v1.3#converttorestiditemid-restversion--string): Converts an item ID formatted for EWS into REST format.
- Added [Office.MailboxEnums.ItemNotificationMessageType](https://dev.office.com/reference/add-ins/outlook/1.3/Office.MailboxEnums?product=outlook&version=v1.3#itemnotificationmessagetype-string): Specifies the notification message type for an appointment or message.
- Added [Office.MailboxEnums.RestVersion](https://dev.office.com/reference/add-ins/outlook/1.3/Office.MailboxEnums?product=outlook&version=v1.3#restversion-string): Specifies the version of the REST API that corresponds to a REST-formatted item ID.
- Added [NotificationMessages](https://dev.office.com/reference/add-ins/outlook/1.3/NotificationMessages?product=outlook&version=v1.3) object: Provides methods for accessing notification messages in an Outlook add-in.
- Added [NotificationMessageDetails](https://dev.office.com/reference/add-ins/outlook/1.3/simple-types?product=outlook&version=v1.3#notificationmessagedetails) type: Returned by the `NotificationMessages.getAllAsync` method.

## Additional resources

- [Outlook add-ins](../../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Get started](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
