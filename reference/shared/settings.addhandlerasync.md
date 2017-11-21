

# Settings.addHandlerAsync method
Adds an event handler for the  **settingsChanged** event.

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**Last changed in**|1.0|

```js
Office.context.document.settings.addHandlerAsync(eventType, handler [, options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](https://dev.office.com/reference/add-ins/shared/eventtype-enumeration)|Specifies the type of event to add. Required.||
| _handler_|**object**|The event handler function to add. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object that you can access from the callback function's only parameter.

In the callback function passed to the  **addHandlerAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.value)|Always returns  **undefined** because there is no data or object to retrieve when adding an event handler.|
|[AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.status)|Determine the success or failure of the operation.|
|[AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.error)|Access an [Error](https://dev.office.com/reference/add-ins/shared/error) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](https://dev.office.com/reference/add-ins/shared/asyncresult.asynccontext)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

You can add multiple event handlers for the specified  _eventType_ as long as the name of each event handler function is unique.


 >**Important**:  Your add-in's code can register a handler for the  **settingsChanged** event when the add-in is running with any Excel client, but the event will fire only when the add-in is loaded with a spreadsheet that is opened in Excel Online, _and_ more than one user is editing the spreadsheet (co-authoring). Therefore, effectively the **settingsChanged** event is supported only in Excel Online in co-authoring scenarios.


## Example




```js
function addSelectionChangedEventHandler() {
    Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, MyHandler);
}

function MyHandler(eventArgs) {
    write('Event raised: ' + eventArgs.type);
    doSomethingWithSettings(eventArgs.settings);
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
|**Excel**||Y||

|||
|:-----|:-----|
|**Available in requirement sets**|Settings|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

