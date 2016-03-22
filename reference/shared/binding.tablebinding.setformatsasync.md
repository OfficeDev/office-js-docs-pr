
# TableBinding.setFormatsAsync method
Sets or updates formatting on specified items and data in the bound table.

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Not in a set|
|**Added in**|1.1|

```
bindingObj.setFormatsAsync(cellFormat [,options] , callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _cellFormat_|**array**|An array that contains one or more JavaScript objects that specify which cells to target and the formatting to apply to them. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **goToByIdAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no data or object to retrieve when setting formats.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

 **Specifying the cellFormat parameter**

Use the  _cellFormat_ parameter to set or change cell formatting values, such as width, height, font, background, alignment, and so on. The value you pass as the _cellFormat_ parameter is an **array** that contains a list of one or more JavaScript objects that specify which cells to target ( `cells:`) and the formats ( `format:`) to apply to them.

Each JavaScript object in the  _cellFormat_ array has this form:

 `{cells:{` _cell_range_ `}, format:{` _format_definition_ `}}`

The  `cells:` property specifies the range you want format using one of the following values:


**Supported ranges in cells property**


|**cells range settings**|**Description**|
|:-----|:-----|
| `{row: i}`|Specifies the range that extends to the ith row of data in the table.|
| `{column: i}`|Specifies the range that extends to ith column of data in the table.|
| `{row: i, column: j}`|Specifies the range of cells from the ith row to the jth column of data in the table.|
| `Office.Table.All`|Specifies the entire table, including column headers, data, and totals (if any).|
| `Office.Table.Data`|Specifies only the data in the table (no headers and totals).|
| `Office.Table.Headers`|Specifies only the header row.|


The  `format:` property specifies values that correspond to a subset of the settings available in the **Format Cells** dialog box in Excel (Right-click > **Format Cells** or **Home** > **Format** > **Format Cells**).

You specify the value of the  `format:` property as a list of one or more _property name_ - _value_ pairs in a JavaScript object literal. The _property name_ specifies the name of the formatting property to set, and _value_ specifies the property value. You can specify multiple values for a given format, such as both a font's color and size. Here's three `format:` property value examples:




```
//Set cells: font color to green and size to 15 points.
format: {fontColor : "green", fontSize : 15}
```




```
//Set cells: border to dotted blue.
format: {borderStyle: "dotted", borderColor: "blue"}
```




```
//Set cells: background to red and alignment to centered.
format: {backgroundColor: "red", alignHorizontal: "center"}
```

You can specify number formats by specifying the number formatting "code" string in the  `numberFormat:` property. The number format strings you can specify correspond to those you can set in Excel using the **Custom** category on the **Number** tab of the **Format Cells** dialog box. This example shows how to format a number as a percentage with two decimal places:




```
format: {numberFormat:"0.00%"}
```

For more detail, see how to [create a custom number format](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1#BM1).



 **Specifying a single target**

The following example shows a  _cellFormat_ value that sets the font color of the header row to red.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: Office.Table.Headers, format: {fontColor: "red"}}], 
    function (asyncResult){});
```

 **Specifying multiple targets**

The  **setFormatsAsync** method can support formatting multiple targets within the bound table in a single function call. To do that, you pass a list of objects in the _cellFormat_ array for each target that you want to format. For example, the following line of code will set the font color of the first row yellow, and the fourth cell in the third row to have a white border and bold text.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

To set formatting on tables when writing data, use the  _tableOptions_ and _cellFormat_ optional parameters of the [Document.setSelectedDataAsync](http://msdn.microsoft.com/library/4c1e13e9-b61a-47df-836c-3ca9aba4ca1c%28Office.15%29.aspx) or [TableBinding.setDataAsync](http://msdn.microsoft.com/library/5b6ecf6f-c57f-4c0d-9605-59daee8fde13%28Office.15%29.aspx) methods.

Setting formatting with the optional parameters of the  **Document.setSelectedDataAsync** and **TableBinding.setDataAsync** methods only works to set formatting when writing data the first time. To make formatting changes after writing data, use the following methods:


- To update cell formatting, such as font color and style, use the  **TableBinding.setFormatsAsync** method (this method).
    
- To update table options, such as banded rows and filter buttons, use the [TableBinding.setTableOptions](../../reference/shared/binding.tablebinding.settableoptionsasync.md) method.
    
- To clear formatting, use the [TableBinding.clearFormats](../../reference/shared/binding.tablebinding.clearformatsasync.md) method.
    
 **Additional remarks for Excel Online**

The number of  _formatting groups_ passed to the _cellFormat_ parameter can't exceed 100. A single formatting group consists of a set of formatting applied to a specified range of cells. For example, the following call passes two formatting groups to _cellFormat_.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});

```

For more details and examples, see [How to format tables in add-ins for Excel](../../docs/excel/format-tables-in-add-ins-for-excel.md).


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**||**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Not in a set.|
|**Minimum permission level**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel in Office for iPad.|
|1.1|Introduced|
