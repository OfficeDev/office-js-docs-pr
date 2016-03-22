
# TableBinding object
Represents a binding in two dimensions of rows and columns, optionally with headers.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Last changed in Selection**|1.1|

```
TableBinding
```


## Members


**Properties**


|**Name**|**Description**|**Updates for Office.js v1.1**|
|:-----|:-----|:-----|
|[columnCount](../../reference/shared/binding.tablebinding.columncount.md)|Gets the number of columns in the specified  **TableBinding** object.|Added support for table binding in content add-ins for Access.|
|[hasHeaders](../../reference/shared/binding.tablebinding.hasheaders.md)|If the specified  **TableBinding** has headers, returns true; otherwise false.|Added support for table binding in content add-ins for Access.|
|[rowCount](../../reference/shared/binding.tablebinding.rowcount.md)|The number of rows in the specified  **TableBinding** object.|For performance reasons, always returns -1 in content add-ins for Access.|

**Methods**


|**Name**|**Description**|**Updates for Office.js v1.1**|
|:-----|:-----|:-----|
|[addColumnsAsync](../../reference/shared/binding.tablebinding.addcolumnsasync.md)|Adds columns and values to a table.||
|[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|Adds rows and values to a table.|Added support for table binding in content add-ins for Access.|
|[clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md)|Clears formatting on the bound table.|New in Office.js v1.1 for add-ins for Excel.|
|[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|Deletes all non-header rows and their values in the table, shifting appropriately for the host application.|Added support for table binding in content add-ins for Access.|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|Writes data to the bound section of the document represented by the specified binding object.|<ul><li>Added support for table binding in content add-ins for Access.</li><li>Added support for setting formatting when writing data to bound tables in add-ins for Excel.</li></ul>|
|[setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|Sets cell and table formatting on specified items and data in the bound table.|Can set table formatting in add-ins for Excel.|
|[setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md)|Updates table formatting options on the bound table.|Can set table formatting in add-ins for Excel.|

## Remarks

The  **TableBinding** object inherits the [id](../../reference/shared/binding.id.md) property, [type](../../reference/shared/binding.type.md) property, [getDataAsync](../../reference/shared/binding.getdataasync.md) method, and [setDataAsync](../../reference/shared/binding.setdataasync.md) method from the [Binding](../../reference/shared/binding.md) abstract object.

After you establish a table binding in Excel, each new row a user adds to the table is automatically included in the binding ( **rowCount** will increase).


## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|TableBindings|
|**Minimum permission level**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for [setting formatting when inserting tables](../../docs/excel/format-tables-in-add-ins-for-excel.md) in Excel.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
