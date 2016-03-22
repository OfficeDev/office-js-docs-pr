
# ActiveView enumeration
Specifies the state of the active view of the document, for example, whether the user can edit the document.

|||
|:-----|:-----|
|**Introduced in Office.js version**|1.1|

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Added in**|1.1|



```
Office.ActiveView
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.ActiveView.Read|"read"|The active view of the host application only lets the user read the content in the document.|
|Office.ActiveView.Edit|"edit"|The active view of the host application lets the user edit the content in the document.|

## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

|||
|:-----|:-----|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint in Office for iPad.|
|1.1|Introduced|
