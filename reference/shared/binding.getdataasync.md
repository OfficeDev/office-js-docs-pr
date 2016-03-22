
# Binding.getDataAsync method
Returns the data contained within the binding.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement sets](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Last changed in TableBindings**|1.1|

```
bindingObj.getDataAsync([, options] , callback );
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|Specifies how to coerce the data being set. ||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|Specifies whether values, such as numbers and dates, are returned with their formatting applied.||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|Specifies if a filter must be applied when the data is retrieved.||
| _rows_|**Office.TableRange.ThisRow**| Specifies the pre-defined string "thisRow" to get data in the currently selected row.|Only for table bindings in content add-ins for Access.|
| _startRow_|**number**|For table or matrix bindings, specifies the zero-based starting row for a subset of the data in the binding. ||
| _startColumn_|**number**|For table or matrix bindings, specifies the zero-based starting column for a subset of the data in the binding. ||
| _rowCount_|**number**|For table or matrix bindings, specifies the number of rows offset from the  _startRow_. ||
| _columnCount_|**number**|For table or matrix bindings, specifies the number of columns offset from the  _startColumn_.||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **Binding.getDataAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Access the values in the specified binding.If the  _coercionType_ parameter is specified (and the call is successful), the data is returned in the format described in the [CoercionType](../../reference/shared/coerciontype-enumeration.md) enumeration topic.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

If an optional parameter is omitted, the following default value is used (when applicable to type and format of the data).



|**Parameter**|**Default**|
|:-----|:-----|
| _coercionType_|The original, uncoerced type of the binding.|
| _valueFormat_|Unformatted data.|
| _filterType_|All values (not filtered).|
| _startRow_|The first row.|
| _startColumn_|The first column.|
| _rowCount_|All rows.|
| _columnCount_|All columns.|
When called from a [MatrixBinding](../../reference/shared/binding.matrixbinding.md) or [TableBinding](../../reference/shared/binding.tablebinding.md), the  **getDataAsync** method will return a subset of the bound values if the optional _startRow_,  _startColumn_,  _rowCount_, and  _columnCount_ parameters are specified (and they specify a contiguous and valid range).


## Example




```
function showBindingData() {
    Office.select("bindings#MyBinding").getDataAsync(function (asyncResult) {
        write(asyncResult.value)
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



There is an important difference in behavior between using the  `"table"` and `"matrix"` _coercionType_ with the **Binding.getDataAsync** method, with respect to data formatted with header rows, as shown in the following two examples. These code examples show event handler functions for the [Binding.SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) event.

If you specify the  `"table"` _coercionType_, the [TableData.rows](../../reference/shared/tabledata.rows.md) property ( `result.value.rows` in the following code example) returns an array that contains only the body rows of the table. So, its 0th row will be the first non-header row in the table.




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'table', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value.rows[0][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
}
```

However, if you specify the  `"matrix"` _coercionType_,  `result.value` in the following code example returns an array that contains the table header in the 0th row. If the table header contains multiple rows, then these are all included in the `result.value` matrix as separate rows before the table body rows are included.




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'matrix', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value[1][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
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
|**Available in requirement sets**|MatrixBindings, TableBindings, TextBindings|
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for table bindings in add-ins for Access.|
|1.0|Introduced|

## See also



#### Other resources


[Bind to regions in a document or spreadsheet](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
