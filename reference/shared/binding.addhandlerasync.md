
# Binding.addHandlerAsync method
Adds a handler to the binding for the specified event type.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|BindingEvents|
|**Last changed in**|1.1|

```
bindingObj.addHandlerAsync(eventType, handler [, options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|Specifies the type of event to add. Required.For a  **Binding** object event, the _eventType_ parameter can be specified as **Office.EventType.BindingSelectionChanged**,  **Office.EventType.BindingDataChanged**, or the corresponding text values of these enumerations.||
| _handler_|**object**|The event handler function to add.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **addHandlerAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no data or object to retrieve when adding an event handler.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

You can add multiple event handlers for the specified  _eventType_ as long as the name of each event handler function is unique.


## Example

The following code sample calls the [select](../../reference/shared/office.select.md) method of the **Office** object to access the binding with ID "MyBinding", and then calls the **addHandlerAsync** method to add a handler function for the [bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md) event of that binding.


```js
function addEventHandlerToBinding() {
    Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
}

function onBindingDataChanged(eventArgs) {
    write("Data has changed in binding: " + eventArgs.binding.id);
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
|**Available in requirement sets**|BindingEvents|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
