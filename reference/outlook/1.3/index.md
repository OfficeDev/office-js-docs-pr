# Outlook add-in API requirement set 1.3

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a [requirement set](../tutorial-api-requirement-sets.md) other than the latest requirement set. 

## What's new in 1.3?

Requirement set 1.3 includes all of the features of [Requirement set 1.2](../1.2/index.md). It added the following features.

- Added support for [add-in commands](../../docs/outlook/add-in-commands-for-outlook.md).
- Added ability to save or close an item being composed.
- Enhanced [Body](Body.md) object to allow addins to get or set the entire body.
- Added conversion methods to convert IDs between EWS and REST formats.
- Added ability to add notification messages to the info bar on items.

### Change log

- Added [Body.getAsync](Body.md#getAsync): Returns the current body in a specified format.
- Added [Body.setAsync](Body.md#setAsync): Replaces the entire body with the specified text.
- Added [Office.context.officeTheme](Office.context.md#officeTheme): Provides access to the Office theme colors.
- Added [Event](Event.md) object: Passed as a parameter to UI-less command functions in an Outlook add-in. Used to signal completion of processing.
- Added [Office.context.mailbox.item.close](Office.context.mailbox.item.md#close): Closes the current item that is being composed.
- Added [Office.context.mailbox.item.saveAsync](Office.context.mailbox.item.md#saveAsync): Asynchronously saves an item.
- Added [Office.context.mailbox.item.notificationMessages](Office.context.mailbox.item.md#notificationMessages): Gets the notification messages for an item.
- Added [Office.context.mailbox.convertToEwsId](Office.context.mailbox.md#convertToEwsId): Converts an item ID formatted for REST into EWS format.
- Added [Office.context.mailbox.convertToRestId](Office.context.mailbox.md#convertToRestId): Converts an item ID formatted for EWS into REST format.
- Added [Office.MailboxEnums.ItemNotificationMessageType](Office.MailboxEnums.md#ItemNotificationMessageType): Specifies the notification message type for an appointment or message.
- Added [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#RestVersion): Specifies the version of the REST API that corresponds to a REST-formatted item ID.
- Added [NotificationMessages](NotificationMessages.md) object: Provides methods for accessing notification messages in an Outlook add-in.
- Added [NotificationMessageDetails](simple-types.md#NotificationMessageDetails) type: Returned by the `NotificationMessages.getAllAsync` method.

## Additional resources

- [Outlook add-ins](../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Get started](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
