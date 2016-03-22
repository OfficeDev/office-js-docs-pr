
# BindingSelectionChangedEventArgs object
Provides information about the binding that raised the [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) event.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Last changed in TableBinding**|1.1|

```
Office.EventType.BindingSelectionChanged
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[binding](../../reference/shared/binding.bindingselectionchangedevent.binding.md)|Gets a [Binding](../../reference/shared/binding.md) object that represents the binding that raised the **SelectionChanged** event.|
|[columnCount](../../reference/shared/binding.bindingselectionchangedevent.columncount.md)|Gets the number of columns selected.|
|[rowCount](../../reference/shared/binding.bindingselectionchangedevent.rowcount.md)|Gets the number of rows selected.|
|[startRow](../../reference/shared/binding.bindingselectionchangedevent.startrow.md)|Gets the index of the first row of the selection (zero-based).|
|[startColumn](../../reference/shared/binding.bindingselectionchangedevent.startcolumn.md)|Gets the index of the first column of the selection (zero-based).|
|[type](../../reference/shared/binding.bindingselectionchangedevent.type.md)|Gets an [EventType](../../reference/shared/eventtype-enumeration.md) enumeration value that identifies the kind of event that was raised.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
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
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for table binding in add-ins for Access.|
|1.0|Introduced|
