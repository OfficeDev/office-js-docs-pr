
# Document.ActiveViewChanged event
Occurs when the user changes the current view of the document.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Introduced in**|1.1|

```
Office.EventType.ActiveViewChanged
```


## Remarks

To add an event handler for the  **ActiveViewChanged** event of a document, use the [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) method of the **Document** object. The event handler receives an argument of type [ActiveViewChangedEventArgs](https://dev.office.com/reference/add-ins/shared/document.activeviewchangedeventargs).


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for Mac**|**Office for iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y||Y|Y|

>**Note: This event will not fire in PowerPoint Online scenarios, as SlideShow mode is treated as a new session. To get the active view, you must query for it during Office.Initialize.
 

|||
|:-----|:-----|
|**Introduced in**|1.1|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|
