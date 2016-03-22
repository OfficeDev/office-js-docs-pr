
# Slice.data property
Gets the raw data of the file slice.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|File|
|**Last changed in**|1.1|

```
var sliceData = slice.data;
```


## Return value

The raw data of the file slice in  **Office.FileType.Text** ("text") or **Office.FileType.Compressed** ("compressed") format as specified by the _fileType_ parameter of the call to the [Document.getFileAsync](../../reference/shared/document.getfileasync.md) method.


## Remarks

Files in the "compressed" format will return a byte array that can be transformed to a base64-encoded string if required.


## Support details


A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|


|||
|:-----|:-----|
|**Available in requirement sets**|File|
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint and Word in Office for iPad.|
|1.0|Introduced|
