
# PlatformType enumeration
Specifies the OS or other platform on which the Office host application is running.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word, PowerPoint, Outlook, OneNote, Project, Access|
|**Last changed**|1.1|

```js
Office.PlatformType
```

## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.PlatformType.PC|"pc"|The platform is PC (Windows).|
|Office.PlatformType.OfficeOnline|"officeOnline"|The platform is Office Online.|
|Office.PlatformType.Mac|"mac"|The platform is Mac.|
|Office.PlatformType.iOS|"iOS"|The platform an iOS device.|
|Office.PlatformType.Android|"android"|The platform is an Android device.|
|Office.PlatformType.Universal|"universal"|The platform is WinRT.|



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
