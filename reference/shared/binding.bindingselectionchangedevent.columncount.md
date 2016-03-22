
# BindingSelectionChangedEventArgs.columnCount property
Gets the number of columns selected.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Last changed in**|1.1|

```
var colCount = eventArgsObj.columnCount;
```


## Return value

The number of columns selected. If a single cell is selected returns 1.


## Remarks

If the user makes a non-contiguous selection, the count for the last contiguous selection within the binding is returned. 

For Word, this property will work only for bindings of [BindingType](../../reference/shared/bindingtype-enumeration.md) "table". If the binding is of type "matrix", **null** is returned. Also, the call will fail if the table contains merged cells, because the structure of the table must be uniform for this property to work correctly.


## Example

The following example adds an event handler for the [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) event to the binding with an [id](../../reference/shared/binding.id.md) of `myTable`. When the user changes the selection, the handler displays the coordinates of the first cell in the selection, and the number of row and columns selected.


```js
function addSelectionHandler() {
    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addHandlerAsync("bindingSelectionChanged", myHandler);
    });
}

// Display selection start coordinates and row/column count.
function myHandler(bArgs) {
    write("Selection start row/col: " + bArgs.startRow + "," + bArgs.startColumn);
    write("Selection row count: " + bArgs.rowCount);
    write("Selection col count: " + bArgs.columnCount);
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


| |**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history





****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|You can now add and remove event handlers for the  **SelectionChanged** event in content add-ins for Access.|
|1.0|Introduced|
