
# TableBinding.addRowsAsync method
Adds rows and values to a table.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Last changed in**|1.1|

```js
bindingObj.addRowsAsync(rows, [,options], callback);
```


## Parameters

_rows_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type: **Array**

&nbsp;&nbsp;&nbsp;&nbsp;An array of arrays that contains one or more rows of data to add to the table. Required.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type: **object**

&nbsp;&nbsp;&nbsp;&nbsp;Specifies the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).
    
&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type: **array, boolean, null, number, object, string, or undefined**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;A user-defined item of any type that is returned in the [AsyncResult](../../reference/shared/asyncresult.md) object without being altered. Optional.<br/><br/>

_callback_<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type:  **object**
    
&nbsp;&nbsp;&nbsp;&nbsp;A function that is invoked when the callback returns, whose only parameter is of type [AsyncResult](../../reference/shared/asyncresult.md). Optional.



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _rows_|**array**|An array of arrays that contains one or more rows of data to add to the table. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **addRowsAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no object or data to retrieve.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

The success or failure of an  **addRowsAsync** operation is atomic. That is, the entire add rows operation must succeed, or it will be completely rolled back (and the **AsyncResult.status** property returned to the callback will report failure):


- Each row in the array you pass as the  _data_ argument must have the same number of columns as the table being updated. If not, the entire operation will fail.
    
- Each row and cell in the array must successfully add that row and cell to the table in the newly added row(s). If any row or cell fails to be set for any reason, the entire operation will fail.
    
 **Additional remarks for Excel Online**

The total number of cells in the value passed to the  _rows_ parameter can't exceed 20,000 in a single call to this method.


## Example




```js
function addRowsToTable() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        var binding = asyncResult.value;
        binding.addRowsAsync([["6", "k"], ["7", "j"]]);
    });
}

```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|TableBindings|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad|
|1.1|Added support for writing table data in add-ins for Access.|
|1.0|Introduced|
