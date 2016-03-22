
# BindingDataChangedEventArgs.binding property
Gets a  **Binding** object that represents the binding that raised the [DataChanged](../../reference/shared/binding.bindingdatachangedevent.md) event.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Last changed in BindingEvents**|1.1|

```js
var myBinding = eventArgsObj.binding;
```


## Return Value

A [Binding](../../reference/shared/binding.md) object.


## Support details


A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
