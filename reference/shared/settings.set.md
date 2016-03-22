

# Settings.set method
Sets or creates the specified setting.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**Last changed in**|1.1|

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Update+a+Row+in+a+Table)


```js
Office.context.document.settings.set(name, value);
```


## Parameters



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **string**

&nbsp;&nbsp;&nbsp;&nbsp;The case-sensitive name of the setting to set or create.

    
_value_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **string**,  **number**,  **boolean**,  **null**,  **object** or **array**

&nbsp;&nbsp;&nbsp;&nbsp;Specifies the value to be stored.
    

## Remarks

The  **set** method creates a new setting of the specified name if it does not already exist, or sets an existing setting of the specified name in the in-memory copy of the settings property bag. After you call the [Settings.saveAsync](../../reference/shared/settings.saveasync.md) method, the value is stored in the document as the serialized JSON representation of its data type. A maximum of 2MB is available for the settings of each add-in.


 >**Important**:  Be aware that the  **Settings.set** method affects only the in-memory copy of the settings property bag. To make sure that additions or changes to settings will be available to your add-in the next time the document is opened, at some point after calling the **Settings.set** method and before the add-in is closed, you must call the **Settings.saveAsync** method to persist settings in the document.


## Example




```js
function setMySetting() {
    Office.context.document.settings.set('mySetting', 'mySetting value');
}

```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

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
|1.1|Added support for custom settings in content add-ins for Access.|
|1.0|Introduced|
