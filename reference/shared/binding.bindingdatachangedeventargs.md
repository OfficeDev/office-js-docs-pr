
# BindingDataChangedEventArgs object
Provides information about the binding that raised the [DataChanged](../../reference/shared/binding.bindingdatachangedevent.md) event.

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
|[binding](../../reference/shared/binding.bindingdatachangedeventargs.binding.md)|Gets a [Binding](../../reference/shared/binding.md) object that represents the binding that raised the **DataChanged** event.|
|[type](../../reference/shared/binding.bindingdatachangedeventargs.type.md)|Gets an [EventType](../../reference/shared/eventtype-enumeration.md) enumeration value that identifies the kind of event that was raised.|

## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

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




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for this event in add-ins for Access.|
|1.0|Introduced|
