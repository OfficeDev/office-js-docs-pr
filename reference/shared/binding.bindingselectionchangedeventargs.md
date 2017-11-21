
# BindingSelectionChangedEventArgs object
Provides information about the binding that raised the [SelectionChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent) event.

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
|[binding](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent.binding)|Gets a [Binding](https://dev.office.com/reference/add-ins/shared/binding) object that represents the binding that raised the **SelectionChanged** event.|
|[columnCount](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent.columncount)|Gets the number of columns selected.|
|[rowCount](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent.rowcount)|Gets the number of rows selected.|
|[startRow](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent.startrow)|Gets the index of the first row of the selection (zero-based).|
|[startColumn](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent.startcolumn)|Gets the index of the first column of the selection (zero-based).|
|[type](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent.type)|Gets an [EventType](https://dev.office.com/reference/add-ins/shared/eventtype-enumeration) enumeration value that identifies the kind of event that was raised.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

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
