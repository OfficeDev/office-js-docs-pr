
# Context.displayLanguage property
Gets the locale (language) specified by the user for the UI of the Office host application.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```
var myDisplayLanguage = Office.context.displayLanguage;
```


## Return Value

A  **string** in the RFC 1766 Language tag format, such as `en-US`.


## Remarks

The  **displayLanguage** value reflects the current **Display Language** setting specified with **File** > **Options** > **Language** in the Office host application.

In content add-ins for Access web apps, the  **displayLanguage** property gets the add-in language (e.g., "en-US").


## Example




```js
function sayHelloWithDisplayLanguage() {
    var myDisplayLanguage = Office.context.displayLanguage;
    switch (myDisplayLanguage) {
        case 'en-US':
            write('Hello!');
            break;
        case 'en-NZ':
            write('G\'day mate!');
            break;
    }
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y|||
|**Excel**|Y|Y|Y||
|**Outlook**|Y|Y||Y|
|**PowerPoint**|Y|Y|Y||
|**Project**|Y||||
|**Word**|Y|Y|Y||

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
|1.1|Added access to this API in content add-ins for Access.|
|1.0|Introduced|
