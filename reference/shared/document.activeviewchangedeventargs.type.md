
# DocumentActiveViewChangedEventArgs.type property
Gets an  **EventType** enumeration value that identifies the kind of event that was raised.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Added in**|1.1|

```
var myEventType = eventArgsObj.type;
```


## Return Value

The [EventType](../../reference/shared/eventtype-enumeration.md) of the event that was raised.


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
|1.1|Introduced.|
