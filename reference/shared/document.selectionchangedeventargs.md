
# DocumentSelectionChangedEventArgs object
Provides information about the document that raised the [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) event.

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Word|
|**Added in**|1.1|

```

```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[document](../../reference/shared/document.selectionchangedeventargs.document.md)|Gets a  **Document** object that represents the document that raised the **SelectionChanged** event.|
|[type](../../reference/shared/document.selectionchangedeventargs.type.md)|Get an  **EventType** enumeration value that identifies the kind of event that was raised.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.0|Introduced|
