

# Office.select method
Creates a promise to return a binding based on the selector string passed in.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement sets](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Last changed in**|1.1|

```js
Office.select(str, onError);
```


## Parameters


_str_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type: **string**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;The selector string to parse and create a promise for.

_onError_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type: **function**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**. Optional.
    

## Callback Value

When the function you passed to the  _onError_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter. If the operation failed, use the [AsyncResult.error](../../reference/shared/asyncresult.error.md) property to access an [Error](../../reference/shared/error.md) object that provides information about the error.


## Remarks

The  **Office.select** method provides access to a [Binding](../../reference/shared/binding.md) object promise that attempts to return the specified binding when any of its asynchronous methods are invoked.

Supported formats: "bindings# _bindingId_", which returns a  **Binding** object for the binding with the [id](../../reference/shared/binding.id.md) of `bindingId`. For more information, see [Asynchronous programming in Office Add-ins](../../docs/develop/asynchronous-programming-in-office-add-ins.md#asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings) and [Bind to regions in a document or spreadsheet](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


 >**Note**: If the  **select** method promise successfully returns a **Binding** object, that object exposes only the following four methods of the [Binding](../../reference/shared/binding.md) object: [getDataAsync](../../reference/shared/binding.getdataasync.md), [setDataAsync](../../reference/shared/binding.setdataasync.md), [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md), and [removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md). If the promise cannot return a  **Binding** object, the _onError_ callback can be used to access an [asyncResult.error](../../reference/shared/asyncresult.error.md) object to get more information.If you need to call a member of the  **Binding** object other than the four methods exposed by the **Binding** object promise returned by the **select** method, instead use the [getByIdAsync](../../reference/shared/bindings.getbyidasync.md) method by using the [Document.bindings](../../reference/shared/document.bindings.md) property and [Bindings.getByIdAsync](../../reference/shared/bindings.getbyidasync.md) method to retrieve the **Binding** object.


## Example

The following code example uses the  **select** method to retrieve a binding with the **id** " `cities`" from the  **Bindings** collection, and then calls the [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) method to add an event handler for the [dataChanged](../../reference/shared/binding.bindingdatachangedevent.md) event of the binding.


```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}
```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Minimum permission level**|[ReadDocument (ReadAllDocument for Open Office XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad|
|1.1|Added the use of the  **select** method to return table bindings created in content add-ins for Access.|
|1.0|Introduced|
