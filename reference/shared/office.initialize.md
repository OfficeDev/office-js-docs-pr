
# Office.initialize event
Occurs when the runtime environment is loaded and the add-in is ready to start interacting with the application and hosted document. 

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```js
Office.initialize = function (reason) {/* initialization code */}
```


## Remarks

The  _reason_ parameter of the **initialize** event listener function returns an [InitializationReason](../../reference/shared/initializationreason-enumeration.md) enumeration value that specifies how initialization occurred. A task pane or content add-in can be initialized in two ways:


- The user just inserted it from  **Recently Used Add-ins** section of the **Add-in** drop-down list on the **Insert** tab of the ribbon in the Office host application, or from **Insert add-in** dialog box.
    
- The user opened a document that already contains the add-in.
    

 >**Note**: The reason parameter of the  **initialize** event listener function only returns an **InitializationReason** enumeration value for task pane and content add-ins. It does not return a value for Outlook add-ins.


## Example

You can use the value of the  **InitializationEnumeration** to implement different logic for when the add-in is first inserted versus when it is already part of the document. The following example shows some simple logic that uses the value of the _reason_ parameter to display how the task pane or content add-in was initialized.


```js
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Display initialization reason.
    if (reason == "inserted")
    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details


A capital Y in the following matrix indicates that this event is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this event.

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
|**Add-in types**|Content, Outlook, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for initializing content add-ins for Access.|
|1.0|Introduced|
