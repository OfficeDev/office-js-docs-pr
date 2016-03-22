
# Error object
Provides specific information about an error that occurred during an asynchronous data operation.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```
asyncResult.error
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[code](../../reference/shared/error.code.md)|Gets the numeric code of the error.|
|[name](../../reference/shared/error.name.md)|Gets the name of the error.|
|[message](../../reference/shared/error.message.md)|Gets a detailed description of the error.|

## Remarks

The  **Error** object is accessed from the [AsyncResult](../../reference/shared/asyncresult.md) object that is returned in the function passed as the _callback_ argument of an asynchronous data operation, such as the [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) method of the **Document** object.


## Example

The following example uses the  **setSelectedDataAsync** method to set the selected text to "Hello World!", and if that fails, displays the values of the **name** and **message** properties of the **Error** object.


```js
function setText() {

    Office.context.document.setSelectedDataAsync("Hello World!", {},
        function (asyncResult) {
            if (asyncResult.status === "failed")
            var err = asyncResult.error; 
                write(err.name + ": " + err.message);
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
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for content add-ins for Access.|
|1.0|Introduced|
