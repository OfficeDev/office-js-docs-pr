
# File object
Represents the document file associated with an Office Add-in.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|File|
|**Last changed in**|1.1|

```
file
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|**[size](https://dev.office.com/reference/add-ins/shared/file.size)**|Gets the document file size in bytes.|
|**[sliceCount](https://dev.office.com/reference/add-ins/shared/file.slicecount)**|Gets the number of slices into which the file is divided.|

**Methods**


|**Name**|**Description**|
|:-----|:-----|
|**[closeAsync](https://dev.office.com/reference/add-ins/shared/file.closeasync)**|Closes the document file.|
|**[getSliceAsync](https://dev.office.com/reference/add-ins/shared/file.getsliceasync)**|Returns the specified slice.|

## Remarks

Access the  **File** object with the [AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.value) property in the callback function passed to the [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync) method.


## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


|||||
|:-----|:-----|:-----|:-----|
||Office for Windows desktop|Office Online (in browser)|Office for iPad|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement set**|File|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint and Word in Office for iPad.|
|1.0|Introduced|
