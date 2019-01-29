---
title: Outlook add-in API Preview requirement set
description: ''
ms.date: 10/31/2018
localization_priority: Priority
---

# Outlook add-in API Preview requirement set

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest. Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.

The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## Features in preview

The following features are in preview.

- [AttachmentContent](/javascript/api/outlook/office.attachmentcontent) - Added a new object that represents the content of an attachment.
- [InternetHeaders](/javascript/api/outlook/office.internetheaders) - Added a new object that represents the internet headers of a message item.
- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) - A new optional parameter `options`, which is a dictionary with one valid value `allowEvent`. This value is used to cancel execution of an event.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.
- [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) - Added a new method to get the content of a specific attachment.
- [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) - Added a new method that gets an item's attachments in compose mode.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - Added a new method that gets an object which represents the sharedProperties of an appointment or message item.
- [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) - Added a new property that represents the internet headers on a message item.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.
- [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) - Added a new enum that specifies the formatting that applies to an attachment's content.
- [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus) - Added a new enum that specifies whether an attachment was added to or removed from an item.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - Added a new bit flag enum that specifies the delegate permissions.
- [Office.EventType](/javascript/api/office/office.eventtype) - Modified to support AttachmentsChanged and OfficeThemeChanged events through addition of `AttachmentsChanged` and `OfficeThemeChanged` entries respectively.
- [SupportsSharedFolders manifest element](../../manifest/supportssharedfolders.md) - Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element. It defines whether the add-in is available in delegate scenarios.

## See also

- [Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](https://docs.microsoft.com/outlook/add-ins/quick-start)
