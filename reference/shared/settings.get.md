
# Settings.get method
Retrieves the specified setting.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**Last changed in**|1.1|

```js
var mySetting = Office.context.document.settings.get(name);
```


## Parameters



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **string**

&nbsp;&nbsp;&nbsp;&nbsp;The case-sensitive name of the setting to retrieve.

    



## Return Value

An  **object** that has property names mapped to JSON serialized values.


## Example




```js
function displayMySetting() {
    write('Current value for mySetting: ' + Office.context.document.settings.get('mySetting'));
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details


A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Settings|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support to create settings in content add-ins for Access.|
|1.0|Introduced|
