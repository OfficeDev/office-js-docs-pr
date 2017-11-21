
# File.getSliceAsync method
Returns the specified slice.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|File|
|**Added in**|1.0|

```js
File.getSliceAsync(sliceIndex, callback);
```


## Parameters


_sliceIndex_ <br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **number**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Specifies the zero-based index of the slice to be retrieved. Required.<br/><br/>
    
_callback_ <br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **object**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;A function that is invoked when the callback returns, whose only parameter is of type [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult). Optional.
    

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getSliceAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.value)|Access the [Slice](https://dev.office.com/reference/add-ins/shared/slice) object.|
|[AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.status)|Determine the success or failure of the operation.|
|[AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.error)|Access an [Error](https://dev.office.com/reference/add-ins/shared/error) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](https://dev.office.com/reference/add-ins/shared/asyncresult.asynccontext)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement set**|File|
|**Minimum permission level**|[ReadDocument (ReadAllDocument required to get Office OpenXML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint and Word in Office for iPad.|
|1.0|Introduced|
