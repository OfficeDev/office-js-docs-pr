

# Office.useShortNamespace method
Toggles on and off the  `Office` alias for the full `Microsoft.Office.WebExtension` namespace.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```js
Office.useShortNamespace(useShortcut);
```


## Parameters



_useShortcut_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **boolean**

    
&nbsp;&nbsp;&nbsp;&nbsp;**true** to use the shortcut alias; otherwise **false** to disable it. The default is **true**.
    


## Example



```js
function startUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(true);
    }
    else {
        Office.useShortNamespace(true);
    }
    write('Office alias is now ' + typeof Office);
}

function stopUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(false);
    }
    else {
        Office.useShortNamespace(false);
    }
    write('Office alias is now ' + typeof Office);
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
|**Add-in types**|Content, Outlook, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for calling this method in content add-ins for Access.|
|1.0|Introduced|
