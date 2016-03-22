
# TableBinding.setTableOptionsAsync method
Updates table formatting options on the bound table.

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Not in a set|
|**Added in**|1.1|

```
bindingObj.setTableOptionsAsync(tableOptions [,options] , callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _tableOptions_|**object**|An object literal containing a list of property name-value pairs that define the table options to apply. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **goToByIdAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no data or object to retrieve when setting table options.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Example

The following example shows how to:


-  **Create an object literal** that specifies the [table formatting options](../../docs/excel/format-tables-in-add-ins-for-excel.md) to update on the bound table.
    
-  **Call setTableOptions** on a previously bound table (with an **id** of `myBinding`) passing the object with formatting settin as the  _tableOptions_ parameter.
    

```js
function updateTableFormatting(){
    var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

    Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
}
```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Not in a set.|
|**Minimum permission level**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel in Office for iPad.|
|1.1|Introduced|
