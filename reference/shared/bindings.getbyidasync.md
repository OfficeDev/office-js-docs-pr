
# Bindings.getByIdAsync method
Gets the specified binding by its identifier.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts, MatrixBindings, TableBindings, TextBindings|
|**Last changed** in|1.1|

```
bindingsObj.getByIdAsync(id [, options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _id_|**string**|Specifies the unique name to be used to identify the binding object. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getByIdAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.value)|Access the [Binding](https://dev.office.com/reference/add-ins/shared/binding) object specified by the _id_ in the call.|
|[AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.status)|Determine the success or failure of the operation.|
|[AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.error)|Access an [Error](https://dev.office.com/reference/add-ins/shared/error) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](https://dev.office.com/reference/add-ins/shared/asyncresult.asynccontext)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

Fails if the specified  _id_ does not exist.


## Example




```js
function displayBindingType() {
    Office.context.document.bindings.getByIdAsync('MyBinding', function (asyncResult) {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
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


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|CustomXmlParts, MatrixBindings, TableBindings, TextBindings|
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for table binding in content add-ins for Access. |
|1.0|Introduced|
