
# Document.setSelectedDataAsync method
Writes data to the current selection in the document.

|||
|:-----|:-----|
|**Hosts:** Access, Excel, PowerPoint, Project, Word, Word Online|**Add-in types: ** Content, Task pane|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Last changed in**|1.1|

```js
Office.context.document.setSelectedDataAsync(data [, options], callback(asyncResult));
```


## Parameters

|Name       | Type  | Description
|:----------|:------|:-----
| data      |object | The data can be any supported [coercision types](#coerciontype)
| options   |object | Specifies a set of [optional parameters](#options)
| callback  |object | [AsyncResult](../../reference/shared/asyncresult.md) object 



## Options
```js
{
    coercionType: '',
    tableOptions: [],
    cellFormat: [],
    imageLeft: 0,
    imageTop: 0,
    imageWidth: 0,
    imageHeight: 0,
    asyncContext
}
```

### coercionType
The follow coercion types are supported by Office.js. Note that not all coercion types are supported by all hosts. 

|Name                       |Access |Excel  |Word   |PowerPoint
|:--------------------------|:-----:|:-----:|:-----:|:---------:|
|Office.CoercionType.Text   |       |   X   |   X   |   X       |
|Office.CoercionType.Matrix |       |   X   |   X   |           |
|Office.CoercionType.Table  |   X   |   X   |   X   |           |
|Office.CoercionType.Html   |       |       |   X   |           |
|Office.CoercionType.Ooxml  |       |       |   X   |           |
|Office.CoercionType.Image  |       |   X   |   X   |   X       |

### tableOptions (object)
For the inserted table, a list of key-value pairs that specify table formatting options, such as header row, total row, and banded rows. (added in 1.1)

### cellFormat (object)
For the inserted table, a list of key-value pairs that specify a range of columns, rows, or cells and the cell formatting to apply to that range. (added in 1.1)

### imageLeft (number)
This option is applicable for inserting images. Indicates the insert location in relation to the left side of the slide for PowerPoint, and its relation to the currently selected cell in Excel. This value is ignored for Word. This value is in points.

### imageTop (number)
This option is applicable for inserting images. Indicates the insert location in relation to the top  of the slide for PowerPoint, and its relation to the currently selected cell in Excel. This value is ignored for Word. This value is in points.

### imageWidth (number)
This option is applicable for inserting images. Indicates the image width. If this option is provided without the imageHeight, the image will scale to match the value of the image width. If both image width and image height are provided, the image will be resized accordingly. If neither the image height or width is provided, then the default image size and aspect ratio will be used. This value is in points.

### imageHeight (number)
This option is applicable for inserting images. Indicates the image height. If this option is provided without the imageWidth, the image will scale to match the value of the image height. If both image width and image height are provided, the image will be resized accordingly. If neither the image height or width is provided, then the default image size and aspect ratio will be used. This value is in points.

### asyncContext (object | value)
A user-defined object that is available on the AsyncResult object's asyncContext property. Use this to provide an object or value to the AsyncResult  when the callback is a named function.


## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **setSelectedDataAsync** method, the [AsyncResult.value](../../reference/shared/asyncresult.value.md) property always returns **undefined** because there is no object or data to retrieve.


## Remarks

The value passed for  _data_ contains the data to write to the current selection. If the value is:


-  **A string:** Plain text or anything that can be coerced to a **string** will be inserted.
    
    
    
    In Excel, you can also specify  _data_ as a valid formula to add that formula to the selected cell. For example, setting _data_ to `"=SUM(A1:A5)"` will total the values in the specified range. However, when you set a formula on the bound cell, after doing so, you can't read the added formula (or any pre-existing formula) from the bound cell. If you call the [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) method on the selected cell to read its data, the method can return only the data displayed in the cell (the formula's result).
    
-  **An array of arrays ("matrix"):** Tabular data without headers will be inserted. For example, to write data to three rows in two columns, you can pass an array like this: `[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`. To write a single column of three rows, pass an array like this:  `[["R1C1"], ["R2C1"], ["R3C1"]]`
    
    
    
    In Excel, you can also specify  _data_ as an array of arrays that contains valid formulas to add them to the selected cells. For example if no other data will be overwritten, setting _data_ to `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` will add those two formulas to the selection. Just as when setting a formula on a single cell as "text", you can't read the added formulas (or any pre-existing formulas) after they have been set - you can only read the formulas' results.
    
-  **A [TableData](../../reference/shared/tabledata.md) object:** A table with headers will be inserted.
    
    
    
     **Note:** In Excel, if you specify formulas in the **TableData** object you pass for the _data_ parameter, you might not get the results you expect due to the "calculated columns" feature of Excel, which automatically duplicates formulas within a column. To work around this when you want to write _data_ that contains formulas to a selected table, try specifying the data as an array of arrays (instead of a **TableData** object), and specify the _coercionType_ as **Microsoft.Office.Matrix** or "matrix".
    
 **Application-specific behaviors**

Additionally, the following application-specific actions apply when writing data to a selection.

 **For Word**


- If there is no selection and the insertion point is at a valid location, the specified  _data_ is inserted at the insertion point as follows:
    
      - If  _data_ is a string, the specified text is inserted.
    
  - If  _data_ is an array of arrays ("matrix") or a **TableData** object, a new Word table is inserted.
    
  - If  _data_ is HTML, the specified HTML is inserted.
    
     >**Important**:  If any of the HTML you insert is invalid, Word won't raise an error. Word will insert as much of the HTML as it can and omits any invalid data.
  - If  _data_ is Office Open XML, the specified XML is inserted.
    
  - If  _data_ is a base64 encoded image stream, the specified image is inserted.
    
- If there is a selection, it will be replaced with the specified  _data_ following the same rules as above.
    
-  **Insert images**: Inserted images are placed inline. The **imageLeft** and **imageTop** parameters are ignored. The image aspect ratio is always locked. If only one of the **imageWidth** and **imageHeight** parameter is given, the other value will be automatically scaled to keep the original aspect ratio.
    
 **For Excel**


- If a single cell is selected:
    
      - If  _data_ is a string, the specified text is inserted as the value of the current cell.
    
  - If  _data_ is an array of arrays ("matrix"), the specified set of rows and columns are inserted, if no other data in surrounding cells will be overwritten.
    
  - If  _data_ is a **TableData** object, a new Excel table with the specified set of rows and headers is inserted, if no other data in surrounding cells will be overwritten.
    
- If multiple cells are selected and the shape does not match the shape of  _data_, an error is returned.
    
- If multiple cells are selected and the shape of the selection exactly matches the shape of  _data_, the values of the selected cells are updated based on the values in  _data_.
    
-  **Insert images**: Inserted images are floating. The position **imageLeft** and **imageTop** parameters are relative to currently selected cell(s). Negative **imageLeft** and **imageTop** values are allowed and possibly readjusted by Excel to position the image inside a worksheet. Image aspect ratio is locked unless both **imageWidth** and **imageHeight** parameters are provided. If only one of the **imageWidth** and **imageHeight** parameter is given, the other value will be automatically scaled to keep the original aspect ratio.
    
In all other cases, an error is returned.

 **For Excel Online**

In addition to the behaviors described for Excel above, the following limits apply when writing data in Excel Online. 


- The total number of cells you can write to a worksheet with the  _data_ parameter can't exceed 20,000 in a single call to this method.
    
- The number of  _formatting groups_ passed to the _cellFormat_ parameter can't exceed 100. A single formatting group consists of a set of formatting applied to a specified range of cells. For example, the following call passes two formatting groups to _cellFormat_.
    

```js
  Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```

 **For PowerPoint**

Inserted images are floating. The position  **imageLeft** and **imageTop** parameters are optional but if provided, both should be present. If a single value is provided, it will be ignored. Negative **imageLeft** and **imageTop** values are allowed and can position an image outside of a slide. If no optional parameter is given and slide has a placeholder, the image will replace the placeholder in the slide. Image aspect ratio will be locked unless both **imageWidth** and **imageHeight** parameters are provided. If only one of the **imageWidth** and **imageHeight** parameter is given, the other value will be automatically scaled to keep the original aspect ratio.


## Example

The following example sets the selected text or cell to "Hello World!", and if that fails, displays the value of the [error.message](../../reference/shared/error.message.md) property.


```js
function writeText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                 write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



Specifying the optional  _coercionType_ parameter lets you specify the kind of data you want to write to a selection. The following example writes data as an array of three rows of two columns, specifying the _coercionType_ as `"matrix"` for that data structure, and if that fails, displays the value of the [error.message](../../reference/shared/error.message.md) property.




```js
function writeMatrix() {
    Office.context.document.setSelectedDataAsync([["Red", "Rojo"], ["Green", "Verde"], ["Blue", "Azul"]], {coercionType: Office.CoercionType.Matrix}
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



The following example writes data as a one column table with a header and four rows, specifying the  _coercionType_ as `"table"` for that data structure, and if that fails, displays the value of the [error.message](../../reference/shared/error.message.md) property.




```js
function writeTable() {
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [['Berlin'], ['Roma'], ['Tokyo'], ['Seattle']];

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: Office.CoercionType.Table},
        function (result) {
            var error = result.error
            if (result.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



 In Word if you want to write HTML to the selection, you can specify the _coercionType_ parameter as `"html"` as shown in the following example, which uses HTML `<b>` tags to make "Hello" bold.




```js
function writeHtmlData() {
    Office.context.document.setSelectedDataAsync("<b>Hello</b> World!", {coercionType: Office.CoercionType.Html}, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In Word, PowerPoint, or Excel, if you want to write an image to the selection, you can specify the  _coercionType_ parameter as `"image"` as shown in the following example. Note that imageLeft and imageTop are ignored by Word.




```js
function insertPictureAtSelection(base64EncodedImageStr) {

    Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
       coercionType: Office.CoercionType.Image,
       imageLeft: 50,
       imageTop: 50,
       imageWidth: 100,
       imageHeight: 100
       },
       function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log("Action failed with error: " + asyncResult.error.message);
           }
       });
}
```


## Support details


A checkmark (![Check symbol](../../images/mod_off15_checkmark.png)) in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**

||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|![Check symbol](../../images/mod_off15_checkmark.png)|||
|**Excel**|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|
|**PowerPoint**|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|
|**Word**|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|


|||
|:-----|:-----|
|**Available in requirement sets**|Selection|
|**Minimum permission level**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support Notes
**Changed in:** 1.1.Support for content add-ins for Access requires  **Selection** requirement set 1.1 or later.Support for setting image data requires  **ImageCoercion** requirement set 1.1 or later. To set this for app activation, use:

```xml
<Requirements>
    <Sets DefaultMinVersion="1.1">
        <Set Name="ImageCoercion"/>
    </Sets>
</Requirements>
```

Runtime detection of ImageCoercion capability can be done with the following code:

```javascript
if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {)) {
    // insertViaImageCoercion();
} 
else {
    // insertViaOoxml();
}
```

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|In Word and Word Online, added support for writing data as a base64 encoded image stream.|
|1.1|In Word Online, added support for writing  _data_ as an **array** of arrays (matrix) and **TableData** (table).|
|1.1|In Excel, PowerPoint and Word in Office for iPad, added the same level of support as Excel, PowerPoint and Word on Windows desktop.|
|1.1|In Word Online, added support for writing  _data_ as **string** (text).|
|1.1|Added support for [setting formatting when inserting tables](../../docs/excel/format-tables-in-add-ins-for-excel.md) with add-ins for Excel by using the _tableOptions_ and _cellFormat_ optional parameters.|
|1.1|Added support for writing table data in add-ins for Access.|
|1.0|Introduced|
