
# DocumentActiveViewChangedEventArgs.activeView property
Gets an  **ActiveView** enumeration value that identifies the state of the active view of the document, for example, whether the user can edit the document.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Added in**|1.1|

```
var myView = eventArgsObj.activeView;
```


## Return Value

The [ActiveView](../../reference/shared/activeview-enumeration.md) of the view that raised the event.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

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
|1.1|Added support for PowerPoint in Office for iPad.|
|1.1|Introduced|
