
# EventType enumeration
Specifies the kind of event that was raised. Returned by the  **type** property of an _EventName_**EventArgs** object.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project, Word|
|**Last changed in Selection**|1.1|

```js
Office.EventType
```


## Members


**Values**


|Enumeration|Value|Description|
|:-----|:-----|:-----|
|Office.EventType.ActiveViewChanged|"documentActiveViewChanged"|A [Document.ActiveViewChanged](/reference/shared/document.activeviewchanged.md) event was raised.|
|Office.EventType.DocumentSelectionChanged|"documentSelectionChanged"|A [Document.SelectionChanged](/reference/shared/document.selectionchanged.event.md) event was raised.|
|Office.EventType.BindingSelectionChanged|"bindingSelectionChanged"|A [Binding.BindingSelectionChanged](/reference/shared/binding.bindingselectionchangedevent.md) event was raised.|
|Office.EventType.BindingDataChanged|"bindingDataChanged"|A [Binding.BindingDataChanged](/reference/shared/binding.bindingdatachangedevent.md) event was raised.|
|Office.EventType.DataNodeDeleted|"nodeDeleted"|A [CustomXmlPart.dataNodeDeleted](/reference/shared/customxmlpart.datanodedeleted.event.md) event was raised.|
|Office.EventType.DataNodeInserted|"nodeInserted"|A [CustomXmlPart.dataNodeInserted](/reference/shared/customxmlpart.datanodeinserted.event.md) event was raised.|
|Office.EventType.DataNodeReplaced|"nodeReplaced"|A [CustomXmlPart.dataNodeReplaced](/reference/shared/customxmlpart.datanodereplaced.event.md) event was raised.|
|Office.EventType.SettingsChanged|"settingsChanged"|A [Settings.settingsChanged](/reference/shared/settings.settingschangedevent.md) event was raised.|

## Remarks


 >**Note**:  Add-ins for Project support the  **Office.EventType.ResourceSelectionChanged**,  **Office.EventType.TaskSelectionChanged**, and  **Office.EventType.ViewSelectionChanged** event types.


## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y||
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.1| Added Office.EventType.ActiveViewChanged enumeration for new **Document.ActiveViewChanged** event.|
|1.0|Introduced|
