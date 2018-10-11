# Office common API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Need information about where add-ins are supported by Office host? See [Office Add-in host and platform availability](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).

Looking for the *host-specific* API requirement sets? See the following API requirement sets:
 
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md) (ExcelApi)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md) (WordApi)
- [OneNote JavaScript API requirement sets](onenote-api-requirement-sets.md) (OneNoteApi)
- [Understanding Outlook API requirement sets](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> We no longer recommend that you create and use Access web apps and databases in SharePoint. As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.

## Common API requirement sets

The following table lists the common API requirement sets, the methods in each set, and the Office host applications that support that requirement set. All of these API requirement sets are version 1.1.

|**Requirement set**|**Office host**|**Methods in set**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac|Document.getActiveViewAsync|
| AddInCommands | See [Add-in command requirement sets](add-in-commands-requirement-sets.md). | |
| BindingEvents  | Access Web Apps<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports output to Office Open XML (OOXML) format as a byte array<br>(Office.FileType.Compressed) when using the Document.getFileAsync method.|
| CustomXmlParts    | Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogApi | See [Dialog API requirement sets](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| File  | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | OneNote Online<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports coercion to HTML (Office.CoercionType.Html) when reading and writing data using the Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| IdentityAPI | See [Identity API requirement sets](identity-api-requirement-sets.md). | Auth.getAccessTokenAsync |
| ImageCoercion | Excel<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports conversion to an image (Office.CoercionType.Image) when writing data using the Document.setSelectedDataAsync method.|
| Mailbox   |Outlook for Windows<br>Outlook for web<br>Outlook for Android<br>Outlook for Mac<br>Outlook Web App |See [Understanding Outlook API requirement sets](outlook-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word<br>Word Online<br>Word for iPad<br>Word for Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports coercion to the "matrix" (array of arrays) data structure (Office.CoercionType.Matrix) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| OoxmlCoercion | Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports coercion to Open Office XML (OOXML) format (Office.CoercionType.Ooxml) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| PartialTableBindings  | Access Web Apps||
| PdfFile   | Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports output to PDF format (Office.FileType.Pdf)<br>when using the Document.getFileAsync method.|
| Selection | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Project<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Access Web Apps<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web Apps<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web Apps<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports coercion to the "table" data structure (Office.CoercionType.Table) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| TextBindings  | Excel<br>Excel Online<br>Excel for iPad<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>Excel for iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Project<br>Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports coercion to text format (Office.CoercionType.Text) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods.|
| TextFile  | Word 2013 and later<br>Word 2016 for Mac and later<br>Word Online<br>Word for iPad|Supports output to text format (Office.FileType.Text) when using the Document.getFileAsync method.|

## Methods that aren't part of a requirement set

The following methods in the JavaScript API for Office aren't part of a requirement set. If your add-in requires any of these methods, use the **Methods** and **Method** elements in the add-in's manifest to declare that they are required, or perform the runtime check using an `if` statement. For more information, see [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

|**Method name**|**Office host support**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access web apps, Excel, Excel Online, and Excel for iPad|
|Document.getFilePropertiesAsync|Excel, Excel Online, Excel for iPad, Excel for Mac, PowerPoint, PowerPoint Online, PowerPoint for iPad, PowerPoint for Mac, Word, Word Online, Word for iPad, and Word for Mac|
|Document.getProjectFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013 and Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Excel for iPad, Excel for Mac, PowerPoint, PowerPoint Online, PowerPoint for iPad, PowerPoint for Mac, Word, Word Online, Word for iPad, and Word for Mac|
|Settings.addHandlerAsync|Access web apps, Excel, Excel Online, PowerPoint, PowerPoint Online, Word, and Word Online|
|Settings.refreshAsync|Access web apps, Excel, Excel Online, PowerPoint, PowerPoint Online, Word, and Word Online|
|Settings.removeHandlerAsync|Access web apps, Excel, Excel Online, PowerPoint, PowerPoint Online, Word, and Word Online|
|TableBinding.clearFormatsAsync|Excel, Excel Online, and Excel for Mac|
|TableBinding.setFormatsAsync|Excel, Excel Online, Excel for iPad, and Excel for Mac|
|TableBinding.setTableOptionsAsync|Excel, Excel Online, Excel for iPad, and Excel for Mac|

## See also

- [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
