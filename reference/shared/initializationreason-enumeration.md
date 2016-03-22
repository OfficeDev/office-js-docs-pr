
# InitializationReason Enumeration
Specifies whether the add-in was just inserted or was already contained in the document. 

|||
|:-----|:-----|
|**Hosts:**|Excel, Project, Word|
|**Added in**|1.0|

```
Office.InitializationReason
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.InitializationReason.Inserted|"inserted"|The add-in was just inserted into the document.|
|Office.InitializationReason.DocumentOpened|"documentOpened"|The add-in is already part of the document that was opened.|

## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
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
|1.1|Added support for Excel and Word in Office for iPad.|
|1.0|Introduced|
