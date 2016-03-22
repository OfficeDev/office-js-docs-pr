
# Format tables in add-ins for Excel


This article explains the different features of the formatting API and outlines how to use them. In this release, you can programmatically specify cell formatting and some other options only for tables (not for  **Office.CoercionType.Text** or **Office.CoercionType.Matrix** data structures) and only in Excel add-ins. To set formatting with your add-in:

- The user selects the table (or where to programmatically insert a table), and then your add-in can call the  **Document.setSelectedDataAsync** method on that table to set formatting.

- Or, if the workbook already contains bound tables (or your add-in uses one of the "addFrom" methods of the [Bindings](../../reference/shared/bindings.bindings.md) object to create bound tables when it is initialized), your add-in can call the **Binding.setDataAsync** method on those bound tables to set formatting.
    
>**Important:** To use these new and updated methods to format tables in Excel add-ins, your add-in project must [use or be updated to use Office.js v1.1 or later](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

## Specifying formatting

To specify the formatting you want to set, you create a JavaScript object literal that contains one or more key-values pairs. You can combine a series of formatting settings in a list within the JavaScript object. For example: 


```js
var myFormat = {fontStyle:"bold", width:"autoFit", borderColor:"purple"};
```

To apply the formatting, pass the JavaScript object to one the methods that support formatting data and other features of the table.

You can work with formatting in two ways:


- The first time your add-in writes data to a selection or binding, by specifying the optional  _cellFormat_ or _tableOptions_ parameters in the _options_ object passed to the [Document.setSelectedDataAysnc](../../reference/shared/document.setselecteddataasync.md) or [Binding.setDataAsync](../../reference/shared/binding.setdataasync.md) methods.
    
- After you initially set formatting, you can [clear or update formatting](#FormatTablesInApps_UpdatingClearing) using one of the new methods dedicated to that purpose.
    

## Using optional parameters with data setting methods

For table bindings, you can use the following optional parameters to specify formatting when setting data with either the  **Document.setSelectedData** or **Binding.setDataAsync** methods: _tableOptions_ and _cellFormat_.


### The tableOptions optional parameter

Use the  _tableOptions_ optional parameter to specify default table styles and turn on and off certain table features, such as: **Header Row**,  **Total Row**, and  **Banded Rows**. The value you pass as the  _tableOptions_ parameter is a JavaScript object that contains a list of key-value pairs. For example,


```js
tableOptions: {bandedRows: true, filterButton: false, style:"TableStyleMedium3"};
```


### The cellFormat optional parameter

Use the  _cellFormat_ optional parameter to change cell formatting values, such as width, height, font, background, alignment, and so on. The value you pass as the _cellFormat_ parameter is an array that contains a list of JavaScript objects that specify which cells to target and the formats to apply to them. For example:


```js
cellFormat: 
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: Office.Table.Headers, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}]
```

You can combine multiple  `cells:` and `format:` pairs in the _cellFormat_ array to minimize the number of function calls required to apply formatting.


#### cells

Use  `cells:` to specify the range of columns, rows, and cells you want to apply formatting to.


**Supported ranges in cells values**


|**cells range settings**|**Description**|
|:-----|:-----|
| `{row: i}`|Specifies the range that extends to the ith row of data in the table.|
| `{column: i}`|Specifies the range that extends to ith column of data in the table.|
| `{row: i, column: j}`|Specifies the range of cells from the ith row to the jth column of data in the table.|
| `Office.Table.All`|Specifies the entire table, including column headers, data, and totals (if any).|
| `Office.Table.Data`|Specifies only the data in the table (no headers and totals).|
| `Office.Table.Headers`|Specifies only the header row.|

#### format

Use  `format:` to specify the formatting you want to apply to the range defined with `cells:` as list of JavaScript key-value pairs. For a list of supported values, see [Supported formatting keys and values](../../docs/excel/format-tables-in-add-ins-for-excel.md#FormatTablesInApps_SupportedFormatting).

 **Limits specifying formatting for Excel Online**

When setting formatting in Excel Online, the number of  _formatting groups_ passed to the _cellFormat_ parameter can't exceed 100. A single formatting group consists of a set of formatting applied to a specified range of cells. (In other words, everything specified in one of the `cells:` object literals in the array passed to _cellFormat_.) For example, the following call passes two formatting groups to  _cellFormat_.




```js
Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```


#### Applying optional parameters

In this release, only the  **Document.setSelectedDataAsync** and **TableBinding.setDataAsync** methods support writing data and setting formatting for tables in the same call using the _tableOptions_ and _cellFormat_ optional parameters. In the following examples, the `tableData` value passed to the first parameter of each method (the _data_ parameter) must be a [TableData](../../reference/shared/tabledata.md) object that contains the definition of the table and data to be written.

 **Document.setSelectedDataAsync example**




```js
Office.context.document.setSelectedDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 **TableBinding.setDataAsync example**




```js
Office.select("bindings#myBinding").setDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 >**Note:**: The call to `Office.select("bindings#myBinding")` assumes that a binding named `myBinding` already exists in the worksheet.


## Updating and clearing formatting


When you set formatting with the  _cellFormat_ and _tableOptions_ optional parameters of the **Document.setSelectedDataAsync** or **TableBinding.setDataAsync** methods, they will set formatting only the first time you call them. To update or clear formatting, you must use three new methods of the **TableBinding** object: **setFormatsAsync**,  **setTableOptionsAsync**, and  **clearFormatsAsync**.


### Updating formatting

The [TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md) method is only for updating cell formatting, such as width, height, font, background, and alignment. It takes _cellFormat_ as the required parameter:


```js
Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

The [TableBinding.setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md) method is only for updating table options, such as banded rows and filter buttons. It takes _tableOptions_ as the required parameter:




```js
var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
```


### Clearing formatting

The [TableBinding.clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md) method is for clearing all formatting in the table. It takes the _asyncContext_ optional parameter and an optional callback function:


```js
Office.select("bindings#myBinding").clearFormatsAsync();
```


## Supported formatting keys and values


The following tables list the supported key-value pairs you can pass to the  _cellFormat_ or _tableOptions_ parameters.

For  `format:` values, available settings correspond to a subset of those in the **Format Cells** dialog box (right-click > **Format Cells** or **Format** > **Format Cells** on the **Home** tab of the ribbon). For `tableOptions:` values, settings correspond to those in the **Table Style Options** and **Table Styles** groups on the **Table Tools** |**Design** tab of the ribbon.


 >**Important**:  The methods of the formatting API support only the options and values summarized below. If you specify formatting options or values other than these, the handling behavior is undefined. These undefined handling behaviors aren't necessarily consistent across supported platforms; you shouldn't develop your add-ins based on any of the side effects of these undefined behaviors for any specific platform. However, the undefined handling behaviors shouldn't harm the state and UI of your add-in or the documents they interact with.


**Alignment**


|**Key**|**Values**|**Notes**|
|:-----|:-----|:-----|
| `alignHorizontal:`|"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "center across selection" \| "distributed"|When combined with an indent value, only the following combinations are supported:<br/><br/><ul><li><code>alignHorizontal: "left"</code> and <code>indentLeft: \<value\></code></li></ul><ul><li><code>alignHorizontal: "right"</code> and <code>indentRight: \<value\></code></li></ul><ul><li><code>alignHorizontal: "distributed"</code> and <code>indentDistributed: \<value\></code></li></ul>|
| `alignVertical:`|"top" \| "center" \| "bottom" \| "justify" \| "distributed"||



**Background**


|**Key**|**Values**|**Notes**|
|:-----|:-----|:-----|
| `backgroundColor:`|"none" \| \<All predefined color names\> \| #RRGGBB|Predefined color names:<br/><br/>"black", "blue", "gray", "green", "orange", "pink", "purple", "red", "teal", "turquoise", "violet", "white", "yellow"|



**Border**


|**Key**|**Values**|**Notes**|
|:-----|:-----|:-----|
| `borderStyle:`|"none" \| \<All predefined border style names\>|Predefined border style names:<br/><br/>"dash dot", "dash dot dot", "dashed", "dotted", "double", "hair", "medium dash dot", "medium dash dot dot", "medium dashed", "medium", "slant dash dot", "thick", "thin"<br/><br/>Applies to all borders in the specified range. (Equivalent to specifying border styles using both the  **Outline** and **Inside** presets on the **Border** tab of the **Format Cells** dialog box.)<br/><br/> **Note:** Excel 2013 supports rendering all 13 predefined border styles. However, Excel Online doesn't support every border style. The following table describes the rendering used for each border style when you open the spreadsheet in Excel Online.<br/><br/><table><tr><th>Excel 2013</th><th>Excel Online</th></tr><tr><td>"dash dot"</td><td>dashed (1px)</td></tr><tr><td>"dash dot dot"</td><td>dotted (1px)</td></tr><tr><td>"dashed"</td><td>dotted (1px)</td></tr><tr><td>"dotted"</td><td>dashed (1px)</td></tr><tr><td>"double"</td><td>double (3px)</td></tr><tr><td>"hair"</td><td>solid (1px)</td></tr><tr><td>"medium dash dot"</td><td>dashed (2px)</td></tr><tr><td>"medium dash dot dot"</td><td>dotted (2px)</td></tr><tr><td>"medium dashed"</td><td>dashed (2px)</td></tr><tr><td>"medium"</td><td>solid (2px)</td></tr><tr><td>"slant dash dot"</td><td>dashed (2px)</td></tr><tr><td>"thick"</td><td>solid (3px)</td></tr><tr><td>"thin"</td><td>solid (1px)</td></tr></table>|
| `borderColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|Applies to all borders in the specified range.|
| `borderTopStyle:`|"none" \| \<All predefined border style names\>|Applies to all borders in the specified range.|
| `borderTopColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|Applies to all borders in the specified range.|
| `borderBottomStyle:`|"none" \| \<All predefined border style names\>|Applies to all borders in the specified range.|
| `borderBottomColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|Applies to all borders in the specified range.|
| `borderLeftStyle:`|"none" \| \<All predefined border style names\>|Applies to all borders in the specified range.|
| `borderLeftColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|Applies to all borders in the specified range.|
| `borderRightStyle:`|"none" \| \<All predefined border style names\>|Applies to all borders in the specified range.|
| `borderRightColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|Applies to all borders in the specified range.|
| `borderOutlineStyle:`|"none" \| \<All predefined border style names\>|Applies to all borders in the specified range.|
| `borderOutlineColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|Applies to all borders in the specified range.|
| `borderInlineStyle:`|"none" \| \<All predefined border style names\>|Applies to only to inside borders in the specified range. (Equivalent to specifying border styles using only the  **Inside** preset on the **Border** tab of the **Format Cells** dialog box.)|
| `borderInlineColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|Applies to only to inside borders in the specified range |



**Cell width, height and wrapping**


|**Key**|**Values**|
|:-----|:-----|
| `width:`|"auto fit" \|  **Number**|
| `height:`|"auto fit" \|  **Number**|
| `wrapping:`|**Boolean**|



**Font**


|**Key**|**Values**|**Notes**|
|:-----|:-----|:-----|
| `fontFamily:`|\<All available font names\>|When you set a font in Excel Online, if the font isn't available in the browser, the API will attempt to fall back to the following fonts in this order: Segoe UI, Thonburi, Arial, Verdana, and Microsoft Sans Serif fonts. If none of these fonts are available, the browser's default font is used.|
| `fontStyle:`|"regular" \| "italic" \| "bold" \| "bold italic"|**Note**: At the time of this publication, setting  `fontStyle:` to "italic" and then subsequently setting "bold" (or vice versa) behaves as a union of these two settings. That is, if, for example, you first set "italic" and then later set "bold", the result will be "bold italic". To set either italic or bold _only_ on a range that was previously set to bold or italic, you must first set `fontStyle:"regular"` to clear the previous formatting.|
| `fontSize:`|**Number**||
| `fontUnderlineStyle:`|"none" \| "single" \| "double" \| "single accounting" \| "double accounting"||
| `fontColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB||
| `fontDirection:`|"context" \| "left-to-right" \| "right-to-left"|Excel Online doesn't currently support displaying text in the right-to-left direction. However, if your add-in sets  `fontDirection:` to "right-to-left" when it's running in Excel Online, that formatting setting is saved in the workbook file and displays correctly when the workbook is opened in Desktop Excel.|
| `fontStrikethrough:`|**Boolean**||
| `fontSuperscript:`|**Boolean**||
| `fontSubScript:`|**Boolean**||
| `fontNormal:`|**Boolean**|Sets the font, font style, size, and effects to the normal style. This resets the cell font formatting to default values. Equivalent to selecting the  **Normal font** check box on the **Font** tab of the **Format Cells** dialog box.|



**Indent**


|**Key**|**Values**|**Notes**|
|:-----|:-----|:-----|
| `indentLeft:`|**Number**|When combined with an alignment value, only the following combination is supported:<br/><br/><ul><li><code>alignHorizontal: "left"</code> and <code>indentLeft: \<value\></code></li></ul>|
| `indentRight:`|**Number**|When combined with an alignment value, only the following combination is supported:<br/><br/><ul><li><code>alignHorizontal: "right"</code> and <code>indentRight: \<value\></code></li></ul>|
| `indentDistributed:`|**Number**|When combined with an alignment value, only the following combination is supported:<br/><br/><ul><li><code>alignHorizontal: "distributed"</code> and <code>indentDistributed: \<value\></code></li></ul>|



**Number formatting**


|**Key**|**Values**|**Notes**|
|:-----|:-----|:-----|
| `numberFormat:`|**String**|To specify number formatting, use a custom number format string. For example, to specify two decimal places with a comma as the thousands separator, you'd specify: <br/><br/> `numberFormat:"#,###.00"`<br/><br/>These are the same custom format strings you can [create with the Custom format category on the Number tab in the Format Cells dialog box](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1).<br/><br/> **Tip:** You can see what a format string looks like for a standard category in the **Format Cells** dialog box in Excel with the following steps:<br/><br/><ol><li>Select a standard format category, for example <span class="ui">Currency</span>, from the <b>Category</b> list.</li><li>Set the format's options in the right side of the dialog box.</li><li>Select the <b>Custom</b> category to view the format string at the top of the <b>Type</b> list.</li></ol>|



**Table options**


|**Key**|**Values**|**Notes**|
|:-----|:-----|:-----|
| `style:`|"none" \| \<All predefined table style names\>|Predefined table style names:<br/><br/>"TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7", "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14", "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleLight21", "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6", "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9", "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12", "TableStyleMedium13", "TableStyleMedium14", "TableStyleMedium15", "TableStyleMedium16", "TableStyleMedium17", "TableStyleMedium18", "TableStyleMedium19", "TableStyleMedium20", "TableStyleMedium21", "TableStyleMedium22", "TableStyleMedium23", "TableStyleMedium24", "TableStyleMedium25", "TableStyleMedium26", "TableStyleMedium27", "TableStyleMedium28", "TableStyleDark1", "TableStyleDark2", "TableStyleDark3", "TableStyleDark4", "TableStyleDark5", "TableStyleDark6", "TableStyleDark7", "TableStyleDark8", "TableStyleDark9", "TableStyleDark10", "TableStyleDark11"<br/><br/>To see what a table style looks like, insert a table in Excel, on the  **Table Tools** \| **Design** tab, choose the **Quick Styles** drop-down, and then select a predefined style. The tooltip for the style will correspond to one of the values in the list above.|
| `headerRow:`|**Boolean**||
| `firstColumn:`|**Boolean**||
| `filterButton:`|**Boolean**||
| `totalRow:`|**Boolean**||
| `lastColumn:`|**Boolean**||
| `bandedRows:`|**Boolean**||
| `bandedColumns:`|**Boolean**||
