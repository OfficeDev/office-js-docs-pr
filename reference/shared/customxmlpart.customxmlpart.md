
# CustomXmlPart object
Represents a single  **CustomXMLPart** in a [CustomXMLParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) collection.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Last changed in**|1.1|

```
Office.context.document.customXmlParts.getByIdAsync(id);
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[builtIn](https://dev.office.com/reference/add-ins/shared/customxmlpart.builtin)|Get a value that indicates whether the CustomXMLPart is built-in.|
|[id](https://dev.office.com/reference/add-ins/shared/customxmlpart.id)|Gets the GUID of the CustomXMLPart|
|[namespaceManager](https://dev.office.com/reference/add-ins/shared/customxmlpart.namespacemanager)|Gets the set of namespace prefix mappings (CustomXMLPrefixMappings) used against the current CustomXMLPart.|

**Methods**


|**Name**|**Description**|
|:-----|:-----|
|[addHandlerAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.addhandlerasync)|Asynchronously adds an event handler for a  **CustomXmlPart** object event.|
|[deleteAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.deleteasync)|Asynchronously deletes this custom XML part from the collection.|
|[getNodesAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.getnodesasync)|Asynchronously gets any CustomXmlNodes in this custom XML part which match the specified XPath.|
|[getXmlAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.getxmlasync)|Asynchronously gets the XML inside this custom XML part.|
|[removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.removehandlerasync)|Removes an event handler for a  **CustomXmlPart** object event.|

**Events**


|**Name**|**Description**|
|:-----|:-----|
|[dataNodeDeleted](https://dev.office.com/reference/add-ins/shared/customxmlpart.datanodedeleted.event)|Occurs when a node is deleted.|
|[dataNodeInserted](https://dev.office.com/reference/add-ins/shared/customxmlpart.datanodeinserted.event)|Occurs when a node is inserted.|
|[dataNodeReplaced](https://dev.office.com/reference/add-ins/shared/customxmlpart.datanodereplaced.event)|Occurs when a node is replaced.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|CustomXmlParts|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
