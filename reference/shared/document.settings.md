
# Document.settings property
Gets an object that represents the saved custom settings of the content or task pane add-in for the current document.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Word|
|**Last changed in**|1.1|

```js
var _settings = Office.context.document.settings;
```

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Persist+Settings)

## Return Value

A [Settings](../../reference/shared/document.settings.md) object.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for content add-ins for Access.|
|1.0|Introduced|
