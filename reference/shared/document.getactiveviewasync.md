
# Document.getActiveViewAsync method
 Returns the state of the current view of the presentation (edit or read).

|||
|:-----|:-----|
|**Hosts:** Excel, PowerPoint, Word|**Add-in types:** Content, Task pane|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|ActiveView|
|**Added in ActiveView**|1.1|

```
Office.context.document.getActiveViewAsync([,options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getActiveViewAsync** method, the [AsyncResult.value](../../reference/shared/asyncresult.value.md) property returns the state of the presentation's current view. The value returned can be either `edit` or `read`.  `edit` corresponds to any of the views in which you can edit slides, such as **Normal** or **Outline View**.  `read` corresponds to either **Slide Show** or **Reading View**.


## Remarks

Can trigger an event when the view changes.


## Example

To get the view of the current presentation, you need to write a callback function that returns that value. The following example shows how to:


-  **Pass an anonymous callback function** that returns the view type to the _callback_ parameter of the **getActiveViewAsync** method.
    
-  **Display the value** on the add-in's page.
    

```js
function getFileView() {
    // Get whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage(asyncResult.value);
        }
    });
}
```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|||Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|ActiveView|
|**Added in ActiveView**|1.1|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history





****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Introduced.|
