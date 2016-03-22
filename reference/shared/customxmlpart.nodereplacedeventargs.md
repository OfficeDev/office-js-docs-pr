
# NodeReplacedEventArgs object
Provides information about the replaced node that raised the [nodeReplaced](../../reference/shared/customxmlpart.nodereplaced.event.md) event.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Last changed in**|1.1|

```
NodeReplacedEventArgs
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|Gets whether the replaced node was inserted as part of an undo or redo operation by the user.|
|[newNode](../../reference/shared/customxmlpart.newnode.md)|Gets the new node.|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|Gets the old (replaced) node.|

## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

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
