
# Slice object
Represents a slice of a document file.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|File|
|**Last changed in**|1.1|

```
slice
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|**[data](../../reference/shared/slice.data.md)**|Gets the raw data of the file slice.|
|**[index](../../reference/shared/slice.index.md)**|Gets the index of the file slice.|
|**[size](../../reference/shared/slice.size.md)**|Gets the size of the slice in bytes.|

## Remarks

The  **Slice** object is accessed with the [File.getSliceAsync](../../reference/shared/file.getsliceasync.md) method.


## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

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




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint and Word in Office for iPad.|
|1.0|Introduced|
