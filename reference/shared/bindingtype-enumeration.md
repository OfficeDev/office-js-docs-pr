
# BindingType enumeration
 Specifies the type of the binding object that should be returned.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Last changed**|1.1|

```
Office.BindingType
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.BindingType.Matrix|"matrix"|Tabular data without a header row. Data is returned as an array of arrays, for example in this form: ` [[row1column1, row1column2],[row2column1, row2column2]]`|
|Office.BindingType.Table|"table"|Tabular data with a header row. Data is returned as a [TableData](../../reference/shared/tabledata.md) object.|
|Office.BindingType.Text|"text"|Plain text. Data is returned as a run of characters.|

## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|Y|||
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
|1.1|Added support for binding table data in add-ins for Access.|
|1.0|Introduced.|
