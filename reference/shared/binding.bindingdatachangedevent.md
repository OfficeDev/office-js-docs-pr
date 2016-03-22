
# Binding.bindingDataChanged event
Occurs when data within the binding is changed.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Last changed in BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## Remarks

To add an event handler for the  **BindingDataChanged** event of a binding, use the [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) method of the **Binding** object. The event handler receives an argument of type [BindingDataChangedEventArgs](../../reference/shared/binding.bindingdatachangedeventargs.md).


## Example




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
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history

|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for this event in add-ins for Access.|
|1.0|Introduced|
