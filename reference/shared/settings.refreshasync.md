

# Settings.refreshAsync method
Reads all settings persisted in the document and refreshes the content or task pane add-in's copy of those settings held in memory.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**Last changed in**|1.1|

```js
Office.context.document.settings.refreshAsync(callback);
```


## Parameters

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **object**

&nbsp;&nbsp;&nbsp;&nbsp;A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.

    



## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object that you can access from the callback function's only parameter.

In the callback function passed to the  **refreshAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.value)|Access a [Settings](https://dev.office.com/reference/add-ins/shared/settings) object with the refreshed values.|
|[AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.status)|Determine the success or failure of the operation.|
|[AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.error)|Access an [Error](https://dev.office.com/reference/add-ins/shared/error) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](https://dev.office.com/reference/add-ins/shared/asyncresult.asynccontext)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

This method is useful in Word and PowerPoint coauthoring scenarios when multiple instances of the same add-in are working against the same document. Because each add-in is working against an in-memory copy of the settings loaded from the document at the time the user opened it, the settings values used by each user can get out of sync. This can happen whenever an instance of the add-in calls the [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) method to persist all of that user's settings to the document. Calling the **refreshAsync** method from the event handler for the [settingsChanged](https://dev.office.com/reference/add-ins/shared/settings.settingschangedevent) event of the add-in will refresh the settings values for all users.

The  **refreshAsync** method can be called from add-ins created for Excel, but since it doesn't support coauthoring there is no reason to do so.


## Example




```js
function refreshSettings() {
    Office.context.document.settings.refreshAsync(function (asyncResult) {
        write('Settings refreshed with status: ' + asyncResult.status);
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
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Settings|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for custom settings in content add-ins for Access.|
|1.0|Introduced|
