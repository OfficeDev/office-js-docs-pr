
# TableData object
Represents the data in a table or a [TableBinding](../../reference/shared/binding.tablebinding.md).

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Added in**|1.1|

```
TableData
```

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Set+Formatting)

## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[headers](../../reference/shared/tabledata.headers.md)|Gets or sets the headers in the table.|
|[rows](../../reference/shared/tabledata.rows.md)|Gets or sets the rows in the table.|

## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|TableBindings|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word Online.|
|1.1|Added support for Excel and Word in Office for iPad|
|1.0|Introduced|
