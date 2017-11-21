
# NodeDeletedEventArgs object
Provides information about the deleted node that raised the [dataNodeDeleted](https://dev.office.com/reference/add-ins/shared/customxmlpart.datanodedeleted.event) event.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Added in**|1.1|

```
NodeDeletedEventArgs
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[isUndoRedo](https://dev.office.com/reference/add-ins/shared/customxmlpart.isundoredo)|Gets whether the node was deleted as part of an Undo/Redo action by the user.|
|[oldNextSibling](https://dev.office.com/reference/add-ins/shared/customxmlpart.oldnextsibling)|Gets the former next sibling of the node that was just deleted from the  **CustomXMLPart** object.|
|[oldNode](https://dev.office.com/reference/add-ins/shared/customxmlpart.oldnode)|Gets the node which was just deleted from the  **CustomXmlPart** object.|

## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

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




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
