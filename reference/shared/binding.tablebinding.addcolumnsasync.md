
# TableBinding.addColumnsAsync method
Adds columns and values to a table.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Last changed in**|1.0|

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Update+a+Row+in+a+Table)


```
bindingObj.addColumnsAsync(data [, options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _data_|**array** or [TableData](../../reference/shared/tabledata.md)|An array of arrays ("matrix") or a  **TableData** object that contains one or more rows of data to add to the table. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **addColumnsAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no object or data to retrieve.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

To add one or more columns specifying the values of the data and headers, pass a  **TableData** object as the _data_ parameter. To add one or more columns specifying only the data, pass an array of arrays ("matrix") as the _data_ parameter.

The success or failure of an  **addColumnAsync** operation is atomic. That is, the entire add columns operation must succeed, or it will be completely rolled back (and the **AsyncResult.status** property returned to the callback will report failure):


- Each row in the array you pass as the  _data_ argument must have the same number of rows as the table being updated. If not, the entire operation will fail.
    
- Each row and cell in the array must successfully add that row or cell to the table in the newly added column(s). If any row or cell fails to be set for any reason, the entire operation will fail.
    
- If you pass a  **TableData** object as the data argument, the number of header rows must match that of the table being updated.
    
**Additional remarks for Excel Online**

The total number of cells in the  **TableData** object passed to the _data_ parameter can't exceed 20,000 in a single call to this method.


## Example

The following example adds a single column with three rows to a bound table with the [id](../../reference/shared/binding.id.md) `"myTable"` by passing a **TableData** object as the _data_ argument of the **addColumnsAsync** method. To succeed, the table being updated must have three rows.


```js
// Add a column to a binding of type table by passing a TableData object.
function addColumns() {
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```

The following example adds a single column with three rows to a bound table with the [id](../../reference/shared/binding.id.md) `myTable` by passing an array of arrays ("matrix") as the _data_ argument of the **addColumnsAsync** method. To succeed, the table being updated must have three rows.




```js
// Add a column to a binding of type table by passing an array of arrays.
function addColumns() {
    var myTable = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
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
|1.1|Added support for Excel and Word in Office for iPad.|
|1.0|Introduced|
