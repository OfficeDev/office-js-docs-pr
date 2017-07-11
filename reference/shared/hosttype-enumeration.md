
# HostType enumeration
Specifies the host Office application in which the add-in is running.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word, PowerPoint, Outlook, OneNote, Project, Access|
|**Last changed**|1.1|

```js
Office.HostType
```

## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.HostType.Word|"word"|The Office host is Microsoft Word.|
|Office.HostType.Excel|"excel"|The Office host is Microsoft Excel.|
|Office.HostType.PowerPoint|"powerPoint"|The Office host is Microsoft PowerPoint.|
|Office.HostType.Outlook|"outlook"|The Office host is Microsoft Outlook.|
|Office.HostType.OneNote|"oneNote"|The Office host is Microsoft OneNote.|
|Office.HostType.Project|"project"|The Office host is Microsoft Project.|
|Office.HostType.Access|"access"|The Office host is Microsoft Access.|



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
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**Add-in types**|Content, Outlook (compose mode), task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Introduced|
