
# Context.commerceAllowed property
Gets whether the add-in is running on a platform that allows links to external payment systems.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Last changed in**|1.1|

```
var allowCommerce = Office.context.commerceAllowed;
```


## Return value

Returns **True** if developers can display sell or upgrade UI in the add-in on that platform; otherwise returns **False**.


## Remarks

The iOS App Store doesn't support apps with add-ins that provide links to additional payment systems. However, Office Add-ins running on the Windows desktop or for Office Online in the browser, do allow such links. If you want the UI of your add-in to provide a link to an external payment system on platforms other than iOS, you can use the  **commerceAllowed** property to control when that link is displayed.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for iPad**|
|:-----|:-----|
|**Excel**|Y|
|**PowerPoint**||
|**Word**|Y|

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
|1.1|Introduced.|
