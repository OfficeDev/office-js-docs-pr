
# TableBinding.rowCount property
Gets the number of rows in the table, as an integer value.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Last changed in Selection**|1.1|

```
var rowCount = bindingObj.rowCount;
```


## Return Value

The number of rows in the specified [TableBinding](../../reference/shared/binding.tablebinding.md) object.


## Remarks

When you insert an empty table by selecting a single row in Excel 2013 and Excel Online (using  **Table** on the **Insert** tab), both Office host applications create a single row of headers followed by a single blank row. However, if your add-in's script creates a binding for this newly inserted table (for example, by using the [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) method), and then checks the value of the **rowCount** property, the value returned will differ depending whether the spreadsheet is open in Excel 2013 or Excel Online.


- In Excel on the desktop,  **rowCount** will return 0 (the blank row following the headers is not counted).
    
- In Excel Online,  **rowCount** will return 1 (the blank row following the headers is counted).
    
You can work around this difference in your script by checking if  `rowCount == 1`, and if so, then checking if the row contains all empty strings.

In content add-ins for Access, for performance reasons the  **rowCount** property always returns -1.


## Example




```js
function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Rows: " + asyncResult.value.rowCount);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details


A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this property.

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
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
