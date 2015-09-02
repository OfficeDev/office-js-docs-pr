# Table

Represents an Excel table.

## [Properties](#getter-and-setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|int|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
|name|string|Name of the table.|
|showHeaders|bool|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
|showTotals|bool|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
|style|string|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|columns|[TableColumnCollection](tablecolumncollection.md)|Represents a collection of all the columns in the table. Read-only.|
|rows|[TableRowCollection](tablerowcollection.md)|Represents a collection of all the rows in the table. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the table.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the table.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with header row of the table.|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire table.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with totals row of the table.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

### delete()
Deletes the table.

#### Syntax
```js
tableObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var table = ctx.workbook.tables.getItem(tableName);
table.delete();
ctx.executeAsync();
```


[Back](#methods)

### getDataBodyRange()
Gets the range object associated with the data body of the table.

#### Syntax
```js
tableObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableDataRange = table.getDataBodyRange();
ctx.executeAsync().then(function () {
		Console.log(tableDataRange.address);
});
```

[Back](#methods)

### getHeaderRowRange()
Gets the range object associated with header row of the table.

#### Syntax
```js
tableObject.getHeaderRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableHeaderRange = table.getHeaderRowRange();
ctx.executeAsync().then(function () {
		Console.log(tableHeaderRange.address);
});
```


[Back](#methods)

### getRange()
Gets the range object associated with the entire table.

#### Syntax
```js
tableObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableRange = table.getRange();
ctx.executeAsync().then(function () {
		Console.log(tableRange.address);
});
```


[Back](#methods)

### getTotalRowRange()
Gets the range object associated with totals row of the table.

#### Syntax
```js
tableObject.getTotalRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableTotalsRange = table.getTotalRowRange();
ctx.executeAsync().then(function () {
		Console.log(tableTotalsRange.address);
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

Get a table by name. 

```js
var ctx = new Excel.RequestContext();
var tableName = 'Table1';
var table = ctx.workbook.tables.getItem(tableName);
ctx.executeAsync().then(function () {
		Console.log(table.index);
});
```

Get a table by index.

```js
var ctx = new Excel.RequestContext();
var index = 0;
var table = ctx.workbook.tables.getItemAt(0);
ctx.executeAsync().then(function () {
		Console.log(table.name);
});
```

Set table style. 

```js
var tableName = 'Table1';
var ctx = new Excel.RequestContext();
var table = ctx.workbook.tables.getItem(tableName);
table.name = 'Table1-Renamed';
table.showTotals = false;
table.tableStyle = 'TableStyleMedium2';
table.load(tableStyle);
ctx.executeAsync().then(function () {
		Console.log(table.tableStyle);
});
```
[Back](#properties)
