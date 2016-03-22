
# Compare Outlook add-in support in Outlook for Mac with other Outlook hosts

You can create and run an Outlook add-in the same way in Outlook for Mac as in the other hosts including Outlook for Windows, OWA for Devices, and Outlook Web App, without customizing the JavaScript for each host. The same calls from the add-in to the JavaScript API for Office generally work the same way, except for the areas described in the following table.

 >**Note**  Outlook for Mac supports JavaScript API for Office in Outlook read mode only.

|**Area**|**Outlook for Windows, OWA for Devices, Outlook Web App**|**Outlook for Mac**|
|:-----|:-----|:-----|
|Supported versions of office.js and Office Add-ins manifest schema|All APIs in Office.js and schema v1.1.|<ul><li>Only APIs that are applicable in read mode. An add-in that uses the new and extended APIs in office.js v1.1 can be activated but the APIs for compose mode will not run properly on Outlook for Mac. </li><li>Schema v1.1.</li></ul>|
|Instances of a recurring appointment series|<ul><li>Can get the item ID and other properties of a master appointment or appointment instance of a recurring series. </li><li>Can use [mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md#displayAppointmentForm) to display an instance or the master of a recurring series.</li></ul>|<ul><li>Can get the item ID and other properties of the master appointment, but not those of an instance of a recurring series.</li><li>Can display the master appointment of a recurring series. Without the item ID, cannot display an instance of a recurring series.</li></ul>|
|Recipient type of an appointment attendee|Can use [EmailAddressDetails.recipientType](../../reference/outlook/simple-types.md) to identify the recipient type of an attendee.|**EmailAddressDetails.recipientType** returns **undefined** for appointment attendees.|
|Version string of the host |The format of the version string returned by [diagnostics.hostVersion](../../reference/outlook/Office.context.mailbox.diagnostics.md) depends on the actual type of host. For example:<ul><li>Outlook for Windows: 15.0.4454.1002</li><li>Outlook Web App: 15.0.918.2</li></ul>|An example of the version string returned by  **Diagnostics.hostVersion** on Outlook for Mac: 15.0 (140325)|
|Custom properties of an item|If the network goes down, an add-in can still access cached custom properties.|Because Outlook for Mac does not cache custom properties, if the network goes down, add-ins would not be able to access them.|
|Attachment details|The content type and attachment names in an [AttachmentDetails](../../reference/outlook/Office.context.mailbox.md) object depend on the type of host:<ul><li>A JSON example of <b>AttachmentDetails.contentType</b>: <b>"contentType": "image/x-png"</b>. </li><li><b>AttachmentDetails.name</b> does not contain any filename extension. As an example, if the attachment is a message that has the subject "RE: Summer activity", the JSON object that represents the attachment name would be <b>"name": "RE: Summer activity"</b>.</li></ul>|<ul><li>A JSON example of <b>AttachmentDetails.contentType</b>: <b>"contentType": "image/png"</b></li><li><b>AttachmentDetails.name</b> always includes a filename extension. Attachments that are mail items have a .eml extension, and appointments have a .ics extension. As an example, if an attachment is an email with the subject "RE: Summer activity", the JSON object that represents the attachment name would be <b>"name": "RE: Summer activity.eml"</b>.</li></ul>|
|String representing the time zone in the  **dateTimeCreated** and **dateTimeModified** properties|As an example: Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)|As an example: Thu Mar 13 2014 14:09:11 GMT+0800 (CST)|
|Time accuracy of  **dateTimeCreated** and **dateTimeModified**|If an add-in uses the following code, the accuracy is up to a millisecond.<br/><pre lang="javascript">JSON.stringify(Office.context.mailbox.item, null, 4);</pre>|The accuracy is up to only a second.|

## Additional resources



- [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md)
    
