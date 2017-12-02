# Office common API requirement sets

> **Important**: We no longer recommend that you create and use Access web apps and databases in SharePoint. As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Need information about where add-ins are supported by Office host? See [Office Add-in host and platform availability](https://dev.office.com/add-in-availability).

Looking for the *host-specific* API requirement sets? See the following API sets:
 
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md) (ExcelApi)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md) (WordApi)
- [OneNote JavaScript API requirement sets](onenote-api-requirement-sets.md) (OneNoteApi)
- [Understanding Outlook API requirement sets](../outlook/tutorial-api-requirement-sets.md) (MailBox)

## Common API requirement sets

The following table lists the common API requirement sets, the methods in each set, and the Office host applications that support that requirement set. All of these API requirement sets are version 1.1.


|**Requirement set**|**Office host**|**Methods in set**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint&nbsp;Online|Document.getActiveViewAsync|
| AddInCommands | See [Add-in command requirement sets](add-in-commands-requirement-sets.md). | |
| BindingEvents  | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | PowerPoint<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>Excel Online<br/>PowerPoint Online|Supports output to Office Open XML (OOXML) format as a byte array<br>(Office.FileType.Compressed) when using the Document.getFileAsync method.|
| CustomXmlParts    | Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| Dialog | See [Dialog API requirement sets](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>PowerPoint Online<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| File  | PowerPoint<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Supports coercion to HTML (Office.CoercionType.Html) when reading and writing data using the Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| IdentityAPI | See [Identity API requirement sets](identity-api-requirement-sets.md). | Auth.getAccessTokenAsync |
| ImageCoercion | Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Supports conversion to an image (Office.CoercionType.Image) when writing data using the Document.setSelectedDataAsync method.|
| Mailbox   |Outlook for Windows<br>Outlook for web<br>Outlook for Mac<br>Outlook Web App |See [Understanding Outlook API requirement sets](../outlook/tutorial-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Word<br>Word Online|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Supports coercion to the "matrix" (array of arrays) data structure (Office.CoercionType.Matrix) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| OoxmlCoercion | Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Supports coercion to Open Office XML (OOXML) format (Office.CoercionType.Ooxml) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| PartialTableBindings  | Access Web Apps||
| PdfFile   | PowerPoint<br/>PowerPoint Online<br/>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Supports output to PDF format (Office.FileType.Pdf)<br>when using the Document.getFileAsync method.|
| Selection | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Access Web Apps<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Supports coercion to the "table" data structure (Office.CoercionType.Table) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| TextBindings  | Excel<br>Excel Online<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Supports coercion to text format (Office.CoercionType.Text) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| TextFile  | Word 2013 and later<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>|Supports output to text format (Office.FileType.Text) when using the Document.getFileAsync method.|

## Methods that aren't part of a requirement set

The following methods in the JavaScript API for Office aren't part of a requirement set. If your add-in requires any of these methods, use the **Methods** and **Method** elements in the add-in's manifest to declare that they are required, or perform the runtime check using an `if` statement. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

|**Method name**|**Office host support**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access web apps, Excel, and Excel Online|
|Document.getFilePropertiesAsync|Excel, Excel Online, Word, Word Online, PowerPoint and PowerPoint Online|
|Document.getProjectFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedViewAsync|PowerPoint and PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Word, and PowerPoint|
|Settings.addHandlerAsync|Access web apps, Excel, Excel Online, Word, Word Online, PowerPoint and PowerPoint Online|
|Settings.refreshAsync|Access web apps, Excel, Excel Online, Word, Word Online, PowerPoint, and PowerPoint Online|
|Settings.removeHandlerAsync|Access web apps, Excel, Excel Online, Word, Word Online, PowerPoint and PowerPoint Online|
|TableBinding.clearFormatsAsync|Excel, Excel Online|
|TableBinding.setFormatsAsync|Excel, Excel Online|
|TableBinding.setTableOptionsAsync|Excel, Excel Online|

## See also

- [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md)
