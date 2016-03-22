
# AsyncResult.value property
Gets the payload or content of this asynchronous operation, if any.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```js
var dataValue = asyncResult.value;
```


## Return Value

Returns the value of the request at the time the asynchronous call was made. 


 >**Note**:  What the  **value** property returns for a particular "Async" method varies depending on the purpose and context of that method. To determine what is returned by the **value** property for an "Async" method, refer to the "Callback value" section of the method's topic. For a complete listing of the "Async" methods, see the Remarks section of the [AsyncResult](../../reference/shared/asyncresult.md) object topic.


## Remarks

You access the  **AsyncResult** object in the function passed as the argument to the _callback_ parameter of an "Async" method, such as the [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) and [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) methods of the **Document** object.


## Example




```js
function getData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Table, function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        }
        else {
            write(asyncResult.value);
        }
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

||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|**Office for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
