
# DocumentActiveViewChangedEventArgs object
Provides information about the view that raised the [ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Introduced in**|1.1|



## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[activeView](https://dev.office.com/reference/add-ins/shared/document.activeviewchangedeventargs.activeview)|Gets an  **ActiveView** enumeration value that identifies the state of the active view of the document, for example, whether the user can edit the document.|
|[type](https://dev.office.com/reference/add-ins/shared/document.activeviewchangedeventargs.type)|Get an  **EventType** enumeration value that identifies the kind of event that was raised.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

|||
|:-----|:-----|
|**Introduced in**|1.1|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|
