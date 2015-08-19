# TableColumn

Represents a column in a table.

## [Properties](#getter-and-setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|int|Returns a unique key that identifies the column within the table. Read-only.|
|index|int|Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|
|name|string|Returns the name of the table column. Read-only.|
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the column from the table.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the column.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with the header row of the column.|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire column.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with the totals row of the column.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

### delete()
Deletes the column from the table.

#### Syntax
```js
tableColumnObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
column.delete();
ctx.executeAsync();
```


[Back](#methods)

### getDataBodyRange()
Gets the range object associated with the data body of the column.

#### Syntax
```js
tableColumnObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var dataBodyRange = column.getDataBodyRange();
dataBodyRange.load(address);
ctx.executeAsync().then(function () {
	Console.log(dataBodyRange.address);
});
```


[Back](#methods)

### getHeaderRowRange()
Gets the range object associated with the header row of the column.

#### Syntax
```js
tableColumnObject.getHeaderRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var headerRowRange = columns.getHeaderRowRange();
headerRowRange.load(address);
ctx.executeAsync().then(function () {
	Console.log(headerRowRange.address);
});
```

[Back](#methods)

### getRange()
Gets the range object associated with the entire column.

#### Syntax
```js
tableColumnObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var columnRange = columns.getRange();
columnRange.load(address);
ctx.executeAsync().then(function () {
	Console.log(columnRange.address);
});
```


[Back](#methods)

### getTotalRowRange()
Gets the range object associated with the totals row of the column.

#### Syntax
```js
tableColumnObject.getTotalRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var totalRowRange = columns.getTotalRowRange();
totalRowRange.load(address);
ctx.executeAsync().then(function () {
	Console.log(totalRowRange.address);
});
```


[Back](#methods)

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

#### Examples
```js

```

[Back](#methods)

### Getter and Setter Examples

```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItem(0);
column.load(index);
ctx.executeAsync().then(function () {
	Console.log(column.index);
});
```

```js
var ctx = new Excel.RequestContext();
var tables = ctx.workbook.tables;
var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
column.values = newValues;
column.load(values);
ctx.executeAsync().then(function () {
	Console.log(column.values);
});
```
[Back](#properties)
