
# ValueFormat enumeration
Specifies whether values, such as numbers and dates, returned by the invoked method are returned with their formatting applied.

|||
|:-----|:-----|
|**Hosts:**|Excel, Project, Word|
|**Added in**|1.0|

```
Office.ValueFormat
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.ValueFormat.Formatted|"formatted"|Return formatted data.|
|Office.ValueFormat.Unformatted|"unformatted"|Return unformatted data.|

## Remarks

For example, if the  _valueFormat_ parameter is specified as `"formatted"`, a number formatted as currency, or a date formatted as mm/dd/yy in the host application will have its formatting preserved. If the  _valueFormat_ parameter is specified as `"unformatted"`, a date will be returned in its underlying sequential serial number form.


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



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.0|Introduced|
