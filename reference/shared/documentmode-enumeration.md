
# DocumentMode enumeration
Specifies whether the document in the associated application is read-only or read-write. 

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Project, Word|
|**Added in**|1.1|

```
Office.DocumentMode
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.DocumentMode.ReadOnly|"readOnly"|The document is read-only.|
|Office.DocumentMode.ReadWrite|"readWrite"|The document can be read and written to.|

## Remarks

Returned by the  **mode** property of the [Document](../../reference/shared/document.md) object.


## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
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
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.0|Introduced|
