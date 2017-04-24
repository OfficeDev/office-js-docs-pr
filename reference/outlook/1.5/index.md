# Outlook add-in API requirement set 1.5

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

## What's new in 1.5?

Requirement set 1.5 includes all of the features of [Requirement set 1.4](../1.4/index.md). 

### Change log

- Added [Office.context.ui.displayDialogAsync](../../shared/officeui.displaydialogasync.md): Displays a dialog box in an Office host.
- Added [Office.context.ui.messageParent](../../shared/officeui.messageparent.md): Delivers a message from the dialog box to its parent/opener page.
- Added [Dialog](../../shared/officeui.dialog.md) object: The object that is returned when the `displayDialogAsync` method is called.

## Additional resources

- [Outlook add-ins](../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Get started](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
