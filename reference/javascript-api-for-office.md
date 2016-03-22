
# JavaScript API for Office
The JavaScript API for Office includes objects, methods, properties, events, and enumerations that you can use in your Office Add-ins code.

Learn more about [supported hosts and other requirements](../docs/overview/requirements-for-running-office-add-ins.md).

The  **Microsoft.Office.WebExtension** namespace (which by default is referenced using the alias [Office](../reference/shared/office.md) in code) contains objects you can use to write script that interacts with content in Office documents, worksheets, presentations, mail items, and projects from your Office Add-ins.
## JavaScript API for Office objects


|**Object**|**Supported add-in type**|**Supported host applications**|
|:-----|:-----|:-----|
|[AsyncResult](../reference/shared/asyncresult.md)|Content add-in, Outlook add-in, Task pane add-in|Access, Excel, Outlook, PowerPoint, Project, Word|
|[AttachmentDetails](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[Binding](../reference/shared/binding.md)|Content add-in, Task pane add-in|Access, Excel, Word|
|[Bindings](../reference/shared/bindings.bindings.md)|Content add-in, Task pane add-in|Access, Excel, Word|
|[Body](../reference/outlook/Body.md)|Outlook add-in|Outlook|
|[Contact](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[Context](../reference/shared/office.context.md)|Content add-in, Outlook add-in, Task pane add-in |Access, Excel, Outlook, PowerPoint, Project, Word|
|[CustomProperties](../reference/outlook/CustomProperties.md)|Outlook add-in|Outlook|
|[CustomXmlNode](../reference/shared/customxmlnode.customxmlnode.md)|Task pane add-in|Word|
|[CustomXmlPart](../reference/shared/customxmlpart.customxmlpart.md)|Task pane add-in |Word|
|[CustomXmlParts](../reference/shared/customxmlparts.customxmlparts.md)|Task pane add-in |Word|
|[CustomXmlPrefixMappings](../reference/shared/customxmlprefixmappings.customxmlprefixmappings.md)|Task pane add-in |Word|
|[Diagnostics](http://msdn.microsoft.com/library/8ad6a159-ed07-4b82-8897-a80fd208551b%28Office.15%29.aspx)|Outlook add-in|Outlook|
|[Document](../reference/shared/document.md)|Content add-in, Task pane add-in|Access, Excel, PowerPoint, Project, Word|
|[EmailAddressDetails](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[EmailUser](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[Entities](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[Error](../reference/shared/error.md)|Content add-in, Outlook add-in, Task pane add-in|Access, Excel, Outlook, PowerPoint, Project, Word|
|[File](../reference/shared/file.md)|Task pane add-in|PowerPoint, Word|
|[Item](../reference/outlook/Office.context.mailbox.item.md)|Outlook add-in|Outlook|
|[Location](../reference/outlook/Location.md)|Outlook add-in|Outlook|
|[Mailbox](../reference/outlook/Office.context.mailbox.md)|Outlook add-in|Outlook|
|[MeetingSuggestion](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[MatrixBinding](../reference/shared/binding.matrixbinding.md)|Content add-in, Task pane add-in|Excel, Word|
|[MeetingSuggestion](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[Office](../reference/shared/office.md)|Content add-in, Outlook add-in, Task pane add-in|Access, Excel, Outlook, PowerPoint, Project, Word|
|[PhoneNumber](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[ProjectDocument](../reference/shared/projectdocument.projectdocument.md)|Task pane add-in |Project|
|[Recipients](../reference/outlook/Recipients.md)|Outlook add-in|Outlook|
|[RoamingSettings](../reference/outlook/RoamingSettings.md)|Outlook add-in|Outlook|
|[Settings](../reference/shared/document.settings.md)|Content add-in, Task pane add-in|Access, Excel, PowerPoint, Word|
|[Slice](../reference/shared/slice.md)|Task pane add-in|PowerPoint, Word, Word Online|
|[Subject](../reference/outlook/Subject.md)|Outlook add-in|Outlook|
|[TableBinding](../reference/shared/binding.tablebinding.md)|Content add-in, Task pane add-in|Access, Excel, Word|
|[TableData](../reference/shared/tabledata.md)|Content add-in, Task pane add-in|Access, Excel, Word|
|[TaskSuggestion](../reference/outlook/simple-types.md)|Outlook add-in|Outlook|
|[TextBinding](../reference/shared/binding.textbinding.md)|Content add-in, Task pane add-in|Excel, Word|
|[Time](../reference/outlook/Time.md)|Outlook add-in|Outlook|
|[UserProfile](../reference/outlook/Office.context.mailbox.userProfile.md)|Outlook add-in|Outlook|




## Enumerations

|**Parent topic**|**Supported add-in type**|**Supported host applications**|
|:-----|:-----|:-----|
|[Enumerations](../reference/shared/enumerations.md)|See child enumeration topics for details.|See Requirements in enumeration topic for details.|

## View APIs by add-in type support

To view the JavaScript API for Office organized by the subsets of the API that support each add-in type, see

|**API **|**Description**|
|:-----|:-----|
|[Shared API](../reference/shared/shared-api.md)|The subset of the API that you can use in all three types of Office Add-ins: content, task pane, and Outlook add-ins.|
|[Document API](../reference/shared/document-api.md)|The subset of the API that you can use in the two types of Office Add-ins associated with documents: content and task pane add-ins.|
|[Mailbox API](../reference/outlook/index.md)|The subset of the API that you can use in Outlook add-ins.|

## Supported host applications
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word
