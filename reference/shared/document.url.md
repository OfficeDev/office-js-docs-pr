
# Document.url property
Gets the URL of the document that the host application currently has open.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Project, Word|
|**Last changed in**|1.1|

```
var docUrl = Office.context.document.url;
```


## Return Value

The URL of the document. Returns  **null** if the URL is unavailable.


## Remarks

 **Important:** The **url** property returns information that may contain personally identifiable information (PII) in the name of the document and location where it is stored. If you must store or transmit this information, be sure to do so in an encrypted format.


## Example




```
function displayDocumentUrl() {
    write(Office.context.document.url);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details


A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history





****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for content add-ins for Access.|
|1.0|Introduced|
