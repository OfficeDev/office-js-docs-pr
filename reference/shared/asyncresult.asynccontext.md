
# AsyncResult.asyncContext property
Gets the user-defined item passed to the optional  _asyncContext_ parameter of the invoked method in the same state as it was passed in.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```
var myContext = asynchResult.asyncContext;
```


## Return Value

Returns the user-defined item (which can be of any JavaScript type:  **String**,  **Number**,  **Boolean**,  **Object**,  **Array**,  **Null**, or  **Undefined**) passed to the optional  _asyncContext_ parameter of the invoked method. Returns **Undefined**, if you didn't pass anything to the _asyncContext_ parameter.


## Example




```js
function getDataWithContext() {
    var format = "Your data: ";
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, { asyncContext: format }, showDataWithContext);
}

 function showDataWithContext(asyncResult) {
    write(asyncResult.asyncContext + asyncResult.value);
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


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**||||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
