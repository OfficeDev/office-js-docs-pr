
# Bindings.document property
Gets a  **Document** object that represents the document associated with this set of bindings.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Last changed in**|1.1|

```
var docObj = bindingsObj.document;
```


## Return Value

A [Document](../../reference/shared/bindings.document.md) object.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added access to a  **Document** object that represents the current Access database in content add-ins for Access.|
|1.0|Introduced|
