
# Document.getFilePropertiesAsync method
Gets file properties of the current document.

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Not in a set|
|**Added in**|1.1|

```js
Office.context.document.getFilePropertiesAsync([, options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getFilePropertiesAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.value)|Returns the file's URL to  `asyncResult.value.url`|
|[AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.status)|Determine the success or failure of the operation.|
|[AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.error)|Access an [Error](https://dev.office.com/reference/add-ins/shared/error) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](https://dev.office.com/reference/add-ins/shared/asyncresult.asynccontext)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

In this release of the API, you can only get the file's URL with the  **url** property ( `asyncResult.value.url`).


## Example

To read the URL of the current file, you need to write a callback function that returns the URL. The following example shows how to:


 - **Pass an anonymous callback function** that returns the value of the file's URL to the _callback_ parameter of the **getFilePropertiesAsync** method.
    
 - **Display the value** on the add-in's page.
    

```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            showMessage("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            showMessage(fileUrl);
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
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Not in a set|
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Introduced.|
