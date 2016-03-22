
# TableBinding.columnCount property
Gets the number of columns in the table, as an integer value.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel,Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Last changed in**|1.1|

```js
var colCount = bindingObj.columnCount;
```


## Return Value

The number of columns in the specified [TableBinding](../../reference/shared/binding.tablebinding.md) object.


## Example




```js
function showBindingColumnCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Column: " + asyncResult.value.columnCount);
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
|**Minimum permission level**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
