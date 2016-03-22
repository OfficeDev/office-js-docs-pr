
# AsyncResultStatus enumeration
Specifies the result of an asynchronous call. 

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```
Office.AsyncResultStatus
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.AsyncResultStatus.Succeeded|"succeeded"|The call succeeded.|
|Office.AsyncResultStatus.Failed|"failed"|The call failed.|

## Remarks

Returned by the [status](../../reference/shared/asyncresult.status.md) property of the [AsyncResult](../../reference/shared/asyncresult.md) object.


## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.


For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|**Office for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y||Y|||

|||
|:-----|:-----|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
