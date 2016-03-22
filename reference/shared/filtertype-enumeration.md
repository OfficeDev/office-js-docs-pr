
# FilterType enumeration
Specifies whether filtering from the host application is applied when the data is retrieved.

|||
|:-----|:-----|
|**Hosts:**|Excel, Project, Word|
|**Last changed in**|1.1|

```js
Office.FilterType
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.FilterType.All|"all"|Return all data (not filtered by the host application).|
|Office.FilterType.OnlyVisible|"onlyVisible"|Return only the visible data (as filtered by the host application).|

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
