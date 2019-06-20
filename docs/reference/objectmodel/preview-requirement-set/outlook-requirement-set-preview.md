---
title: Outlook add-in API Preview requirement set
description: ''
ms.date: 06/14/2019
localization_priority: Priority
---

# Outlook add-in API Preview requirement set

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest. Methods and properties that are introduced in this requirement set should be individually tested for availability before using them. You may also need to join the [Office Insider program](https://products.office.com/office-insider).

The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## Features in preview

The following features are in preview.

### Attachments

#### [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

Added a new object that represents the content of an attachment.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

Added a new method to get the content of a specific attachment.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

Added a new method that gets an item's attachments in compose mode.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

Added a new enum that specifies the formatting that applies to an attachment's content.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

Added a new enum that specifies whether an attachment was added to or removed from an item.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

Added `AttachmentsChanged` event to `Item`.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

---

### Block on send

#### [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`. This value is used to cancel execution of an event.

**Available in**: Outlook on the web (classic)

---

### Categories

In Outlook, a user can group messages and appointments by using a category to color-code them. The user defines categories in a master list on their mailbox. They can then apply one or more categories to an item.

> [!NOTE]
> This feature is not supported in Outlook for iOS or Outlook for Android.

#### [Categories](/javascript/api/outlook/office.categories)

Added a new object that represents an item's categories.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [CategoryDetails](/javascript/api/outlook/office.categorydetails)

Added a new object that represents a category's details (its name and associated color).

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [MasterCategories](/javascript/api/outlook/office.mastercategories)

Added a new object that represents the categories master list on a mailbox.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)

Added a new property that represents the categories master list on a mailbox.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.item.categories](/javascript/api/outlook/office.item#categories)

Added a new property that represents the set of categories on an item.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor)

Added a new enum that specifies the colors available to be associated with categories.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

---

### Delegate access

#### [SharedProperties](/javascript/api/outlook/office.sharedproperties)

Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#getitemidasyncoptions-callback)

Added a new method that gets the ID of a saved appointment or message item.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

Added a new method that gets an object which represents the sharedProperties of an appointment or message item.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

Added a new bit flag enum that specifies the delegate permissions.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [SupportsSharedFolders manifest element](../../manifest/supportssharedfolders.md)

Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element. It defines whether the add-in is available in delegate scenarios.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

---

### Enhanced location

#### [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

Added a new object that represents the set of locations on an appointment.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [LocationDetails](/javascript/api/outlook/office.locationdetails)

Added a new object that represents a location. Read only.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

Added a new object that represents the id of a location.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

Added a new property that represents the set of locations on an appointment.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

Added a new enum that specifies an appointment location's type.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

Added `EnhancedLocationsChanged` event to `Item`.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

---

### Integration with actionable messages

#### [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)

---

### Internet headers

#### [InternetHeaders](/javascript/api/outlook/office.internetheaders)

Added a new object that represents the internet headers of a message item.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheaders)

Added a new property that represents the internet headers on a message item.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

---

### Office theme

#### [Office.context.mailbox.officeTheme](/javascript/api/office/office.officetheme)

Added ability to get Office theme.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

#### [Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Added `OfficeThemeChanged` event to `Mailbox`.

**Available in**: Outlook on Windows (connected to Office 365 subscription)

---

### SSO

#### [Office.context.auth.getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.

**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook for Mac (connected to Office 365 subscription), Outlook on the web (new), Outlook on the web (classic)

## See also

- [Outlook add-ins](/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](/outlook/add-ins/quick-start)
