# What's changed in the JavaScript API for Office

The JavaScript API for Office is periodically updated with new and updated objects, methods, properties, events and enumerations to extend the functionality of your Office Add-ins. Use the links below to see the new and updated API members.

To develop add-ins using new API members, you need to [update the JavaScript API for Office files in your project](https://docs.microsoft.com/office/dev/add-ins/develop/update-your-javascript-api-for-office-and-manifest-schema-version).

To view all API members including those that are unchanged from previous updates, see [JavaScript API for Office](javascript-api-for-office.md).

## New and updated APIs

### New and updated objects

|**Object**|**Description**|**Version added or updated**|
|:-----|:-----|:-----|
|`Item`|Updates and additions to:<br><ul><li><p>The `getSelectedDataAsync` and `setSelectedDataAsync` methods to support getting the user's selection and overwriting it in the subject and body  of a message or appointment.</p></li><li><p>The `displayReplyAllForm` and `displayReplyForm` methods to support adding an attachment to the reply form of an appointment.</p></li></ul>|Mailbox 1.2|
|`Item`|Updated to include methods and fields for creating compose mode Outlook add-ins. |1.1|
|`Binding`|Updated to support table binding in content add-ins for Access.|1.1|
|`Bindings`|Updated to support table binding in content add-ins for Access.|1.1|
|`Body`|Added to enable creating and editing the body of a message or appointment in compose mode Outlook add-ins.|1.1|
|`Document`|Updates and additions to: <ul><li><p>Support <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">mode</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings" target="_blank">settings</a>, and <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">url</a> properties in content add-ins for Access.</p></li><li><p>Get the document as PDF with the <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-" target="_blank">getFileAsync</a> method in add-ins for PowerPoint and Word.</p></li><li><p>Get file properties with the <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-" target="_blank">getFileProperties</a> method in add-ins for Excel, PowerPoint, and Word.</p></li><li><p>Navigate to locations and objects within the document with the <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-" target="_blank">goToByIdAsync</a> method in add-ins for Excel and PowerPoint.</p></li><li><p>Get the id, title, and index for selected slides with the <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-" target="_blank">getSelectedDataAsync</a> method (when you specify the new <span class="keyword">Office.CoercionType.SlideRange</span><a href="https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js" target="_blank">coercionType</a> enum) in add-ins for PowerPoint.</p></li></ul>|1.1|
|`Location`|Added to enable setting the location of an appointment in compose mode Outlook add-ins.|1.1|
|`Office`|Updated the select method to support getting bindings in content add-ins for Access.|1.1|
|`Recipients`|Added to enable getting and setting the recipients of a message or appointment in compose mode.|1.1|
|`Settings`|Updated to support creating custom settings in content add-ins for Access.|1.1|
|`Subject`|Added to enable getting and setting the subject of a message or appointment in compose mode Outlook add-ins.|1.1|
|`Time`|Added to enable getting and setting the start and end time of an appointment in compose mode Outlook add-ins.|1.1|

### New and updated enumerations

|**Object**|**Description**|**Version**|
|:-----|:-----|:-----|
|`ActiveView`|Specifies the state of the active view of the document, for example, whether the user can edit the document.Added so that add-ins for PowerPoint can determine if the users is viewing the presentation ( **Slide Show**) or editing slides. |1.1|
|`CoercionType`|Updated with  **Office.CoercionType.SlideRange** to support getting the selected slide range with the **getSelectedDataAsync** method in add-ins for PowerPoint.|1.1|
|`EventType`|Updated to include the new ActiveViewChanged event.|1.1|
|`FileType`|Updated to specify output in PDF format.|1.1|
|`GoToType`|Added to specify the place or object in the document to go to.|1.1|

