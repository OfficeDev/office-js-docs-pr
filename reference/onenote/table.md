# Table Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a table in a OneNote page.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|columnCount|int|Gets the number of columns in the table. Read-only.|
|id|string|Gets the ID of the table. Read-only.|
|rowCount|int|Gets the number of rows in the table. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the Table object. Read-only.|
|rows|[TableRowCollection](tablerowcollection.md)|Gets all of the table rows. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[appendColumn(values: string[])](#appendcolumnvalues-string)|void|Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise the column is empty.|
|[appendRow(values: string[])](#appendrowvalues-string)|[TableRow](tablerow.md)|Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise the row is empty.|
|[deleteColumns(columnIndex: number, columnCount: number)](#deletecolumnscolumnindex-number-columncount-number)|void|Deletes a contiguous run of columns.|
|[deleteRows(rowIndex: number, rowCount: number)](#deleterowsrowindex-number-rowcount-number)|void|Deletes a contiguous run of rows.|
|[getCell(rowIndex: number, cellIndex: number)](#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Gets the table cell at a specified row and column.|
|[hideBorder()](#hideborder)|void|Hides the table's border|
|[insertColumn(index: number, values: string[])](#insertcolumnindex-number-values-string)|void|Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise the column is empty.|
|[insertRow(index: number, values: string[])](#insertrowindex-number-values-string)|[TableRow](tablerow.md)|Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise the row is empty.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[showBorder()](#showborder)|void|Make's the table's border visible|

## Method Details


### appendColumn(values: string[])
Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise the column is empty.

#### Syntax
```js
tableObject.appendColumn(values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|string[]|Optional. Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.|

#### Returns
void

### appendRow(values: string[])
Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise the row is empty.

#### Syntax
```js
tableObject.appendRow(values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|string[]|Optional. Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.|

#### Returns
[TableRow](tablerow.md)

### deleteColumns(columnIndex: number, columnCount: number)
Deletes a contiguous run of columns.

#### Syntax
```js
tableObject.deleteColumns(columnIndex, columnCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|columnIndex|number|The first column to delete.|
|columnCount|number|Optional. Optional. The number of columns to delete. Default 1.|

#### Returns
void

### deleteRows(rowIndex: number, rowCount: number)
Deletes a contiguous run of rows.

#### Syntax
```js
tableObject.deleteRows(rowIndex, rowCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowIndex|number|The first row to delete.|
|rowCount|number|Optional. Optional. The number of rows to delete. Default 1.|

#### Returns
void

### getCell(rowIndex: number, cellIndex: number)
Gets the table cell at a specified row and column.

#### Syntax
```js
tableObject.getCell(rowIndex, cellIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowIndex|number|The index of the row.|
|cellIndex|number|The index of the cell in the row.|

#### Returns
[TableCell](tablecell.md)

### hideBorder()
Hides the table's border

#### Syntax
```js
tableObject.hideBorder();
```

#### Parameters
None

#### Returns
void

### insertColumn(index: number, values: string[])
Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise the column is empty.

#### Syntax
```js
tableObject.insertColumn(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index where the column will be inserted in the table.|
|values|string[]|Optional. Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.|

#### Returns
void

### insertRow(index: number, values: string[])
Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise the row is empty.

#### Syntax
```js
tableObject.insertRow(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index where the row will be inserted in the table.|
|values|string[]|Optional. Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.|

#### Returns
[TableRow](tablerow.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### showBorder()
Make's the table's border visible

#### Syntax
```js
tableObject.showBorder();
```

#### Parameters
None

#### Returns
void
