
# CustomXmlPart.nodeInserted event
Occurs when a node is inserted.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Last changed in**|1.1|

```
Office.EventType.NodeInserted
```


## Remarks

To add an event handler for the  **nodeInserted** event, use the [addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md) method of the **CustomXmlPart** object.


## Example




```js
function addNodeInsertedEvent() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.addHandlerAsync(Office.EventType.NodeInserted, function (eventArgs) {
            write("A node has been inserted.");
        });
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
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|CustomXmlParts|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
