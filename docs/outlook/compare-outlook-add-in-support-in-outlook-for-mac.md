---
title: Compare Outlook add-in support in Outlook for Mac | Microsoft Docs
description: Learn how add-in support in Outlook for Mac compares with other Outlook hosts.
author: jasonjoh
ms.topic: article
ms.technology: office-add-ins
ms.date: 08/09/2017
ms.author: jasonjoh
---

# Compare Outlook add-in support in Outlook for Mac with other Outlook hosts

You can create and run an Outlook add-in the same way in Outlook for Mac as in the other hosts, including Outlook for Windows, OWA for Devices, and Outlook on the web, without customizing the JavaScript for each host. The same calls from the add-in to the JavaScript API for Office generally work the same way, except for the areas described in the following table.

For more information, see [Deploy and install Outlook add-ins for testing](testing-and-tips.md).

| Area | Outlook for Windows, OWA for Devices, Outlook on the web | Outlook for Mac |
|:-----|:-----|:-----|
| Supported versions of office.js and Office Add-ins manifest schema | All APIs in Office.js and schema v1.1. | <ul><li>Supports requirement sets 1.0, 1.1, 1.2, 1.3, and 1.4</li><li>Schema v1.1.</li></ul>**NOTE**: Outlook for Mac does not support saving a meeting. The `saveAsync` method will fail when called from a meeting in compose mode. |
| Instances of a recurring appointment series | <ul><li>Can get the item ID and other properties of a master appointment or appointment instance of a recurring series.</li><li>Can use [mailbox.displayAppointmentForm](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#displayappointmentformitemid) to display an instance or the master of a recurring series.</li></ul> | <ul><li>Can get the item ID and other properties of the master appointment, but not those of an instance of a recurring series.</li><li>Can display the master appointment of a recurring series. Without the item ID, cannot display an instance of a recurring series.</li></ul> |
| Recipient type of an appointment attendee | Can use [EmailAddressDetails.recipientType](https://docs.microsoft.com/javascript/api/outlook_1_5/office.emailaddressdetails#recipienttype) to identify the recipient type of an attendee. | `EmailAddressDetails.recipientType` returns `undefined` for appointment attendees. |
| Version string of the host | The format of the version string returned by [diagnostics.hostVersion](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.diagnostics#hostversion-string) depends on the actual type of host. For example:<ul><li>Outlook for Windows: `15.0.4454.1002`</li><li>Outlook Web App: `15.0.918.2`</li></ul> |An example of the version string returned by `Diagnostics.hostVersion` on Outlook for Mac: `15.0 (140325)` |
| Custom properties of an item | If the network goes down, an add-in can still access cached custom properties. | Because Outlook for Mac does not cache custom properties, if the network goes down, add-ins would not be able to access them. |
| Attachment details | The content type and attachment names in an [AttachmentDetails](https://docs.microsoft.com/javascript/api/outlook_1_5/office.attachmentdetails) object depend on the type of host:<ul><li>A JSON example of `AttachmentDetails.contentType`: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` does not contain any filename extension. As an example, if the attachment is a message that has the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity"`.</li></ul> | <ul><li>A JSON example of `AttachmentDetails.contentType`: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` always includes a filename extension. Attachments that are mail items have a .eml extension, and appointments have a .ics extension. As an example, if an attachment is an email with the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity.eml"`.<p>**NOTE**: If a file is programmatically attached (e.g through an add-in) without an extension then the `AttachmentDetails.name`  will not contain the extension as part of filename.</p></li></ul> |
| String representing the time zone in the `dateTimeCreated` and `dateTimeModified` properties |As an example: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | As an example: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Time accuracy of `dateTimeCreated` and `dateTimeModified` | If an add-in uses the following code, the accuracy is up to a millisecond:<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| The accuracy is up to only a second. |

