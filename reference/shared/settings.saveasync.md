
# Settings.saveAsync method
Persists the in-memory copy of the settings property bag in the document.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**Last changed in**|1.1|

```js
Office.context.document.settings.saveAsync(callback);
```


## Parameters



_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **object**

&nbsp;&nbsp;&nbsp;&nbsp;A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**. Optional.

    



## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **saveAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no object or data to retrieve.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

Any settings previously saved by an add-in are loaded when it is initialized, so during the lifetime of the session you can just use the [set](../../reference/shared/settings.set.md) and [get](../../reference/shared/settings.get.md) methods to work with the in-memory copy of the settings property bag. When you want to persist the settings so that they are available the next time the add-in is used, use the **saveAsync** method.


 >**Note**:  The  **saveAsync** method persists the in-memory settings property bag into the document file; however, the changes to the document file itself are saved only when the user (or **AutoRecover** setting) saves the document to the file system.

The [refreshAsync](../../reference/shared/settings.refreshasync.md) method is only useful in coauthoring scenarios (which are only supported in Word) when other instances of the same add-in might change the settings and those changes should be made available to all instances.


## Example




```js
function persistSettings() {
    Office.context.document.settings.saveAsync(function (asyncResult) {
        write('Settings saved with status: ' + asyncResult.status);
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
