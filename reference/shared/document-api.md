
# Document API


The Document API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in the two types of Office Add-ins associated with documents: content and task pane add-ins.


## Objects





|**Object**|**Description**|**Supported host applications**|
|:-----|:-----|:-----|
|[Binding](../../reference/shared/binding.md)|An abstract class that represents a binding to a section of the document.|<ul><li>Access</li><li>Excel</li><li>Word</li></ul>|
|[Bindings](../../reference/shared/bindings.bindings.md)|Represents the bindings the add-in has within the document.|<ul><li>Access</li><li>Excel</li><li>Word</li></ul>|
|[CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md)|Represents an XML node in a tree in a document.|<ul><li>Word</li></ul>|
|[CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md)|Represents a single  **CustomXMLPart** in a **CustomXMLParts** collection.|<ul><li>Word</li></ul>|
|[CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md)|Represents a collection of  **CustomXMLPart** objects.|<ul><li>Word</li></ul>|
|[CustomXmlPrefixMappings](../../reference/shared/customxmlprefixmappings.customxmlprefixmappings.md)|Represents a collection of custom namespace prefix mappings.|<ul><li>Word</li></ul>|
|[Document](../../reference/shared/document.md)|An abstract class that represents the document the add-in is interacting with.|<ul><li>Access</li><li>Excel</li><li>PowerPoint</li><li>Project</li><li>Word</li></ul>|
|[File](../../reference/shared/file.md)|Represents the document file associated with an Office Add-in.|<ul><li>PowerPoint</li><li>Word</li></ul>|
|[MatrixBinding](../../reference/shared/binding.matrixbinding.md)|Represents a binding in two dimensions of rows and columns. |<ul><li>Excel</li><li>Word</li></ul>|
|[ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)|An abstract class that represents the project document (the active project) with which the Office Add-in interacts.|<ul><li>Project</li></ul>|
|[Settings](../../reference/shared/document.settings.md)|Represents custom settings for a task pane or content add-in that are stored in the host document as name/value pairs.|<ul><li>Access</li><li>Excel</li><li>PowerPoint</li><li>Word</li></ul>|
|[Slice](../../reference/shared/slice.md)|Represents a slice of a document file.|<ul><li>PowerPoint</li><li>Word</li></ul>|
|[TableBinding](../../reference/shared/binding.tablebinding.md)|Represents a binding in two dimensions of rows and columns, optionally with headers.|<ul><li>Access</li><li>Excel</li><li>Word</li></ul>|
|[TableData](../../reference/shared/tabledata.md)|Represents the data in a table or a  **TableBinding**.|<ul><li>Access</li><li>Excel</li><li>Word</li></ul>|
|[TextBinding](../../reference/shared/binding.textbinding.md)|Represents a bound text selection in the document.|<ul><li>Excel</li><li>Word</li></ul>|

## Supported host applications


|||
|:-----|:-----|
|**Supported hosts**|<ul><li>Access</li><li>Excel</li><li>Outlook</li><li>PowerPoint</li><li>Project</li><li>Word</li></ul><br/>See "Supported host applications" in the Objects table for details about support for each object.|
|**Library**|Office.js|
|**Namespace**|Office|

## Additional resources



- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
    
