# Outlook add-in API requirement set 1.4

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a [requirement set](../tutorial-api-requirement-sets.md) other than the latest requirement set.

## What's new in 1.4?

Requirement set 1.4 includes all of the features of [Requirement set 1.3](../1.3/index.md). It added access to the `Office.ui` namespace.

### Change log

- Added [Office.context.ui.displayDialogAsync](../../shared/officeui.displaydialogasync.md): Displays a dialog box in an Office host.
- Added [Office.context.ui.messageParent](../../shared/officeui.messageparent.md): Delivers a message from the dialog box to its parent/opener page.
- Added [Dialog](../../shared/officeui.dialog.md) object: The object that is returned when the `displayDialogAsync` method is called.

## Additional resources

- [Outlook add-ins](../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Get started](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
