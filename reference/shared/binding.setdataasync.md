
# Binding.setDataAsync method
Writes data to the bound section of the document represented by the specified binding object.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement sets](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Last changed in TableBindings**|1.1|

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Update+a+Row+in+a+Table)


```js
bindingObj.setDataAsync(data [, options] ,callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _data_|<table><tr><td><b>string</b></td><td>Excel, Excel Online, Word, and Word Online only</td></tr><tr><td><b>array</b> (array of arrays – "matrix")</td><td>Excel and Word only</td></tr><tr><td><a href="https://msdn.microsoft.com/en-us/library/office/fp161002"><b>TableData</b></a></td><td data-th="Text value">Access, Excel, and Word only</td></tr><tr><td><b>HTML</b></td><td>Word and Word Online only</td></tr><tr><td><b>Office Open XML</b></td><td>Word only</td></tr></table>|The data to be set in the current selection. Required.|**Changed in:** 1.1.Support for content add-ins for Access requires  **TableBinding** requirement set 1.1 or later.|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|Specifies how to coerce the data being set. ||
| _columns_|**array of strings**| Specifies the column names.|**Added in:** v1.1.Only for table bindings in content add-ins for Access.|
| _rows_|**Office.TableRange.ThisRow**|Specifies the pre-defined string "thisRow" to set data in the currently selected row. |**Added in:** v1.1.Only for table bindings in content add-ins for Access.|
| _startColumn_|**number**|Specifies the zero-based starting column for a subset of the data. |Only for table or matrix bindings. If omitted, data is set starting in the first column.|
| _startRow_|**number**|Specifies the zero-based starting row for a subset of the data in the binding. |Only for table or matrix bindings. If omitted, data is set starting in the first row.|
| _tableOptions_|**object**|For the inserted table, a list of key-value pairs that specify [table formatting options](../../docs/excel/format-tables-in-add-ins-for-excel.md), such as header row, total row, and banded rows. |**Added in:** v1.1. **Supported in:** Excel.|
| _cellFormat_|**object**|For the inserted table, a list of key-value pairs that specify a range of columns, rows, or cells and the [cell formatting](../../docs/excel/format-tables-in-add-ins-for-excel.md) to apply to that range.|**Added in** v1.1. **Supported in:** Excel, Excel Online.|
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **setDataAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no object or data to retrieve.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

The value passed for  _data_ contains the data to be written in the binding. The kind of value passed determines what will be written as described in the following table.



|**_data_ value**|**Data written**|
|:-----|:-----|
|A  **string**|Plain text or anything that can be coerced to a  **string** will be written.|
|An array of arrays ("matrix")|Tabular data without headers will be written. For example, to write data to three rows in two columns, you can pass an array like this:  ` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`To write a single column of three rows, pass an array like this:  `[["R1C1"], ["R2C1"], ["R3C1"]]`|
|A [TableData](../../reference/shared/tabledata.md) object|A table with headers will be written.|
Additionally, these application-specific actions apply when writing data to a binding.

 **For Word**, the specified  _data_ is written to the binding as follows:



|**_data_ value**|**Data written**|
|:-----|:-----|
|A  **string**|The specified text is written.|
|An array of arrays ("matrix") or a  **TableData** object|A Word table is written.|
|HTML|The specified HTML is written.
 >**Important**  If any of the HTML you write is invalid, Word will not raise an error. Word will write as much of the HTML as it can and will omit any invalid data.

|
|Office Open XML ("Open XML")|The specified the XML is written.|
 **For Excel**, the specified  _data_ is written to the binding as follows:



|**_data_ value**|**Data written**|
|:-----|:-----|
|A  **string**|The specified text is inserted as the value of the first bound cell.You can also specify a valid formula to add that formula to the bound cell. For example, setting  _data_ to `"=SUM(A1:A5)"` will total the values in the specified range. However, when you set a formula on the bound cell, after doing so, you can't read the added formula (or any pre-existing formula) from the bound cell. If you call the [Binding.getDataAsync](../../reference/shared/binding.getdataasync.md) method on the bound cell to read its data, the method can return only the data displayed in the cell (the formula's result).|
|An array of arrays ("matrix"), and the shape exactly matches the shape of the binding specified|The set of rows and columns are written.You can also specify an array of arrays that contain valid formulas to add them to the bound cells. For example, setting  _data_ to `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` will add those two formulas to a binding that contains two cells. Just as when setting a formula on a single bound cell, you can't read the added formulas (or any pre-existing formulas) from the binding with the **Binding.getDataAsync** method - it returns only the data displayed in the bound cells.|
|A  **TableData** object, and the shape of the table matches the bound table.|The specified set of rows and/or headers are written, if no other data in surrounding cells will be overwritten. **Note:** If you specify formulas in the **TableData** object you pass for the _data_ parameter, you might not get the results you expect due to the "calculated columns" feature of Excel, which automatically duplicates formulas within a column. To work around this when you want to write _data_ that contains formulas to a bound table, try specifying the data as an array of arrays (instead of a **TableData** object), and specify the _coercionType_ as **Microsoft.Office.Matrix** or "matrix".|
 **Additional remarks for Excel Online**


- The total number of cells in the value passed to the  _data_ parameter can't exceed 20,000 in a single call to this method.
    
- The number of  _formatting groups_ passed to the _cellFormat_ parameter can't exceed 100. A single formatting group consists of a set of formatting applied to a specified range of cells. For example, the following call passes two formatting groups to _cellFormat_.
    
```js
  Office.select("bindings#myBinding).setDataAsync([['Berlin'],['Munich'],['Duisburg']],
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});

```

In all other cases, an error is returned.

The  **setDataAsync** method will write data in a subset of a table or matrix binding if the optional _startRow_ and _startColumn_ parameters are specified, and they specify a valid range.


## Example




```js
function setBindingData() {
    Office.select("bindings#MyBinding").setDataAsync('Hello World!', function (asyncResult) { });
}
```

Specifying the optional  _coercionType_ parameter lets you specify the kind of data you want to write to a binding. For example, in Word if you want to write HTML to a text binding, you can specify the _coercionType_ parameter as `"html"` as shown in the following example, which uses HTML `<b>` tags to make "Hello" bold.




```js
function writeHtmlData() {
    Office.select("bindings#myBinding").setDataAsync("<b>Hello</b> World!", {coercionType: "html"}, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In this example, the call to  **setDataAsync** passes the _data_ parameter as an array of arrays (to create a single column of three rows), and specifies the data structure with the _coercionType_ parameter as a `"matrix"`.




```js
function writeBoundDataMatrix() {
    Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],{ coercionType: "matrix" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In the  `writeBoundDataTable` function in this example, the call to **setDataAsync** passes the _data_ parameter as a **TableData** object (to write three columns and three rows), and specifies the data structure with the _coercionType_ parameter as a `"table"`. 

In the  `updateTableData` function, the call to **setDataAsync** again passes the _data_ parameter as a **TableData** object, but as a single column with a new header and three rows, to update the values in the last column of the table created with the `writeBoundDataTable` function. The optional zero-based _startColumn_ parameter is specified as 2 to replace the values in the third column of the table.




```js
function writeBoundDataTable() {
    // Create a TableData object.
    var myTable = new Office.TableData();
    myTable.headers = ['First Name', 'Last Name', 'Grade'];
    myTable.rows = [['Kim', 'Abercrombie', 'A'], ['Junmin','Hao', 'C'],['Toni','Poe','B']];

    // Set myTable in the binding.
    Office.select("bindings#myBinding").setDataAsync(myTable, { coercionType: "table" }, 
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Error: '+ asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}

// Replace last column with different data.
function updateTableData() {
     var newTable = new Office.TableData();
     newTable.headers = ["Gender"];
     newTable.rows = [["M"],["M"],["F"]];
     Office.select("bindings#myBinding").setDataAsync(newTable, { coercionType: "table", startColumn:2 }, 
         function (asyncResult) {
             if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                 write('Error: '+ asyncResult.error.message);
         } else {
            write('Bound data: ' + asyncResult.value);
         }     
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
|**Available in requirement sets**|MatrixBindings, TableBindings, TextBindings|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|<ul><li>In add-ins for Access, added support for writing table data.</li><li>In add-ins for Excel, added support for <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">setting formatting when writing data to a table binding</a> using the <span class="parameter" sdata="paramReference">tableOptions</span> and <span class="parameter" sdata="paramReference">cellFormat</span> optional parameters.</li></ul>|
|1.0|Introduced|

## See also



#### Other resources


[Bind to regions in a document or spreadsheet](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
