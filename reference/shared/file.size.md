
# File.size property
Gets the document file size in bytes.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint,Word|
|**Added in**|1.1|

```
var fileSize = file.size;
```


## Return value

The file size in bytes.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

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
|1.1|Added support for PowerPoint and Word in Office for iPad.|
|1.0|Introduced|
