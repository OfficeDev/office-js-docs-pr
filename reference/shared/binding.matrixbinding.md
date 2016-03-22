
# MatrixBinding object
Represents a binding in two dimensions of rows and columns. 

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings|
|**Last changed in Selection**|1.1|

```
MatrixBinding
```


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[columnCount](../../reference/shared/binding.matrixbinding.columncount.md)|Gets the number of columns in the matrix data structure, as an integer value.|
|[rowCount](../../reference/shared/binding.matrixbinding.rowcount.md)|Gets the number of rows in the matrix data structure, as an integer value.|

## Remarks

The  **MatrixBinding** object inherits the [id](../../reference/shared/binding.id.md) property, [type](../../reference/shared/binding.type.md) property, [getDataAsync](../../reference/shared/binding.getdataasync.md) method, and [setDataAsync](../../reference/shared/binding.setdataasync.md) method from the [Binding](../../reference/shared/binding.md) object.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|MatrixBindings|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.0|Introduced|
