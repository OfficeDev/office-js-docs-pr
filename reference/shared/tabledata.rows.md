
# TableData.rows property
Gets or sets the rows in the table.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Added in**|1.1|

```
var myRows = tableBindingObj.rows;
```


## Return Value

Returns an array of arrays that contains the data in the table. Returns an empty  **array** `[]`, if there are no rows.


## Remarks

To specify rows, you must specify an array of arrays that corresponds to the structure of the table. For example, to specify two rows of  **string** values in a two-column table you would set the **row** property to ` [['a', 'b'], ['c', 'd']]`.

If you specify  **null** for the **rows** property (or leave the property empty when you construct a **TableData** object), the following results occur when your code executes:


- If you insert a new table, a blank row will be inserted.
    
- If you overwrite or update an existing table, the existing rows are not altered.
    

## Example

The following example creates a single-column table with a header and three rows.


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}
```



[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Write+and+Read+a+Table)

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|


|||
|:-----|:-----|
|**Available in requirement sets**|TableBindings|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word Online.|
|1.1|Added support for Excel and Word in Office for iPad|
|1.0|Introduced|
