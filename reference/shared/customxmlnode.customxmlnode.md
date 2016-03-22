
# CustomXmlNode object
Represents an XML node in a tree in a document.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Last changed in**|1.1|

```js
CustomXmlNode
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[baseName](../../reference/shared/customxmlnode.basename.md)|Gets the base name of the node without the namespace prefix, if one exists.|
|[nodeType](../../reference/shared/customxmlnode.nodetype.md)|Gets the type of the  **CustomXMLNode**.|
|[namespaceUri](../../reference/shared/customxmlnode.namespaceuri.md)|Retrieves the string GUID of the  **CustomXMLPart**.|

**Methods**


|**Name**|**Description**|
|:-----|:-----|
|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|Asynchronously gets the nodes as an array of  **CustomXMLNode** objects matching the relative XPath expression.|
|[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|Asynchronously gets the value of the node.|
|[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|Asynchronously gets the XML of the node.|
|[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|Asynchronously sets the value of the node.|
|[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|Asynchronously sets the XML of the node.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

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
