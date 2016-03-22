
# Error.code property
Gets the numeric code of the error.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in Selection**|1.1|

```
var errCode = asyncResult.error.code;
```


## Return value

The error code as a  **number**.


## Remarks

The  **Error** object and its properties are accessed from the [AsyncResult](../../reference/shared/asyncresult.md) object that is returned in the function passed as the _callback_ argument of an asynchronous data operation.


## Example

To cause an error to be thrown, select a table or a matrix, and then call the  `setText` function.


```js
function setText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            if (asyncResult.status === "failed")
                var error = asyncResult.error;
            write(error.name + ": " + error.code + " - " + error.message);
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



||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|**Outlook for Mac**|
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



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for content add-ins for Access.|
|1.0|Introduced|
