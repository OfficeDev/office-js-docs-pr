
# BindingDataChangedEventArgs object
Provides information about the binding that raised the [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Last changed in BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[binding](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedeventargs.binding)|Gets a [Binding](https://dev.office.com/reference/add-ins/shared/binding) object that represents the binding that raised the **DataChanged** event.|
|[type](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedeventargs.type)|Gets an [EventType](https://dev.office.com/reference/add-ins/shared/eventtype-enumeration) enumeration value that identifies the kind of event that was raised.|

## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

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




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for this event in add-ins for Access.|
|1.0|Introduced|
