
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
|Office.EventType.ActiveViewChanged|"documentActiveViewChanged"|A [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event was raised.|
|Office.EventType.DocumentSelectionChanged|"documentSelectionChanged"|A [Document.SelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) event was raised.|
|Office.EventType.BindingSelectionChanged|"bindingSelectionChanged"|A [Binding.BindingSelectionChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent) event was raised.|
|Office.EventType.BindingDataChanged|"bindingDataChanged"|A [Binding.BindingDataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event was raised.|
|Office.EventType.DataNodeDeleted|"nodeDeleted"|A [CustomXmlPart.dataNodeDeleted](https://dev.office.com/reference/add-ins/shared/customxmlpart.datanodedeleted.event) event was raised.|
|Office.EventType.DataNodeInserted|"nodeInserted"|A [CustomXmlPart.dataNodeInserted](https://dev.office.com/reference/add-ins/shared/customxmlpart.datanodeinserted.event) event was raised.|
|Office.EventType.DataNodeReplaced|"nodeReplaced"|A [CustomXmlPart.dataNodeReplaced](https://dev.office.com/reference/add-ins/shared/customxmlpart.datanodereplaced.event) event was raised.|
|Office.EventType.SettingsChanged|"settingsChanged"|A [Settings.settingsChanged](https://dev.office.com/reference/add-ins/shared/settings.settingschangedevent) event was raised.|

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
