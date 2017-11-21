
# What's changed in the JavaScript API for Office
The JavaScript API for Office is periodically updated with new and updated objects, methods, properties, events and enumerations to extend the functionality of your Office Add-ins. Use the links below to see the new and updated API members.

To develop add-ins using new API members, you need to [update the JavaScript API for Office files in your project](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

To view all API members including those that are unchanged from previous updates, see [JavaScript API for Office](../reference/javascript-api-for-office.md).


## New and updated APIs

 **New and updated objects**


|**Object**|**Description**|**Version added or updated**|
|:-----|:-----|:-----|
|[Item](..https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item)|Updates and additions to:<br><ul><li><p>The <a href="..https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> and <a href="..https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> methods to support getting the user's selection and overwriting it in the subject and body  of a message or appointment.</p></li><li><p>The <a href="..https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item#displayReplyAllForm" target="_blank">displayReplyAllForm</a> and <a href="..https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item#displayReplyForm" target="_blank">displayReplyForm</a> methods to support adding an attachment to the reply form of an appointment.</p></li></ul>|Mailbox 1.2|
|[Item](..https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item)|Updated to include methods and fields for creating compose mode Outlook add-ins. |1.1|
|[Binding](..https://dev.office.com/reference/add-ins/shared/binding)|Updated to support table binding in content add-ins for Access.|1.1|
|[Bindings](..https://dev.office.com/reference/add-ins/shared/bindings.bindings)|Updated to support table binding in content add-ins for Access.|1.1|
|[Body](..https://dev.office.com/reference/add-ins/outlook/Body)|Added to enable creating and editing the body of a message or appointment in compose mode Outlook add-ins.|1.1|
|[Document](..https://dev.office.com/reference/add-ins/shared/document)|Updates and additions to: <ul><li><p>Support <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>, <a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a>, and <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> properties in content add-ins for Access.</p></li><li><p>Get the document as PDF with the <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> method in add-ins for PowerPoint and Word.</p></li><li><p>Get file properties with the <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> method in add-ins for Excel, PowerPoint, and Word.</p></li><li><p>Navigate to locations and objects within the document with the <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> method in add-ins for Excel and PowerPoint.</p></li><li><p>Get the id, title, and index for selected slides with the <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> method (when you specify the new <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a> enum) in add-ins for PowerPoint.</p></li></ul>|1.1|
|[Location](..https://dev.office.com/reference/add-ins/outlook/Location)|Added to enable setting the location of an appointment in compose mode Outlook add-ins.|1.1|
|[Office](..https://dev.office.com/reference/add-ins/shared/office)|Updated the select method to support getting bindings in content add-ins for Access.|1.1|
|[Recipients](..https://dev.office.com/reference/add-ins/outlook/Recipients)|Added to enable getting and setting the recipients of a message or appointment in compose mode.|1.1|
|[Settings](..https://dev.office.com/reference/add-ins/shared/document.settings)|Updated to support creating custom settings in content add-ins for Access.|1.1|
|[Subject](..https://dev.office.com/reference/add-ins/outlook/Subject)|Added to enable getting and setting the subject of a message or appointment in compose mode Outlook add-ins.|1.1|
|[Time](..https://dev.office.com/reference/add-ins/outlook/Time)|Added to enable getting and setting the start and end time of an appointment in compose mode Outlook add-ins.|1.1|



**New and updated enumerations**


|**Object**|**Description**|**Version**|
|:-----|:-----|:-----|
|[ActiveView](..https://dev.office.com/reference/add-ins/shared/activeview-enumeration)|Specifies the state of the active view of the document, for example, whether the user can edit the document.Added so that add-ins for PowerPoint can determine if the users is viewing the presentation ( **Slide Show**) or editing slides. |1.1|
|[CoercionType](..https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration)|Updated with  **Office.CoercionType.SlideRange** to support getting the selected slide range with the **getSelectedDataAsync** method in add-ins for PowerPoint.|1.1|
|[EventType](..https://dev.office.com/reference/add-ins/shared/eventtype-enumeration)|Updated to include the new ActiveViewChanged event.|1.1|
|[FileType](..https://dev.office.com/reference/add-ins/shared/filetype-enumeration)|Updated to specify output in PDF format.|1.1|
|[GoToType](..https://dev.office.com/reference/add-ins/shared/gototype-enumeration)|Added to specify the place or object in the document to go to.|1.1|

## Additional resources


- [Office Add-ins API and schema references](../reference/reference.md)
    
- [Office Add-ins](../docs/overview/office-add-ins.md)
    
