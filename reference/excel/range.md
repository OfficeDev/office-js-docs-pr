# Range Object (JavaScript API for Excel)

Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|address|string|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|addressLocal|string|Represents range reference for the specified range in the language of the user. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|cellCount|int|Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|Represents the total number of columns in the range. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnHidden|bool|Represents if all columns of the current range are hidden.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|columnIndex|int|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|Represents the formula in A1-style notation.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|Represents the formula in R1C1-style notation.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|hidden|bool|Represents if all cells of the current range are hidden. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|Represents Excel's number format code for the given cell.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|Returns the total number of rows in the range. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHidden|bool|Represents if all rows of the current range are hidden.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|rowIndex|int|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|object[][]|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|conditionalFormats|[ConditionalFormatCollection](conditionalformatcollection.md)|Collection of ConditionalFormats that intersect the range. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|format|[RangeFormat](rangeformat.md)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[RangeSort](rangesort.md)|Represents the range sort of the current range. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current range. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[calculate()](#calculate)|void|Calculates a range of cells on a worksheet.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|[clear(applyTo: string)](#clearapplyto-string)|void|Clear range values, format, fill, border, etc.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete(shift: string)](#deleteshift-string)|void|Deletes the cells associated with the range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Gets a column contained in the range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsAfter(count: number)](#getcolumnsaftercount-number)|[Range](range.md)|Gets a certain number of columns to the right of the current Range object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsBefore(count: number)](#getcolumnsbeforecount-number)|[Range](range.md)|Gets a certain number of columns to the left of the current Range object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", it's `getEntireColumn` is a range that represents columns "B:E").|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", it's `GetEntireRow` is a range that represents rows "4:11").|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|Gets the range object that represents the rectangular intersection of the given ranges.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersectionOrNullObject(anotherRange: Range or string)](#getintersectionornullobjectanotherrange-range-or-string)|[Range](range.md)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastCell()](#getlastcell)|[Range](range.md)|Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastRow()](#getlastrow)|[Range](range.md)|Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getResizedRange(deltaRows: number, deltaColumns: number)](#getresizedrangedeltarows-number-deltacolumns-number)|[Range](range.md)|Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Gets a row contained in the range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsAbove(count: number)](#getrowsabovecount-number)|[Range](range.md)|Gets a certain number of rows above the current Range object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsBelow(count: number)](#getrowsbelowcount-number)|[Range](range.md)|Gets a certain number of rows below the current Range object.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|Returns the used range of the given range object. If there are no used cells within the range, this function will throw an ItemNotFound error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getVisibleView()](#getvisibleview)|[RangeView](rangeview.md)|Represents the visible rows of the current range.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[merge(across: bool)](#mergeacross-bool)|void|Merge the range cells into one region in the worksheet.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[select()](#select)|void|Selects the specified range in the Excel UI.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[unmerge()](#unmerge)|void|Unmerge the range cells into separate cells.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### calculate()
Calculates a range of cells on a worksheet.

#### Syntax
```js
rangeObject.calculate();
```

#### Parameters
None

#### Returns
void

### clear(applyTo: string)
Clear range values, format, fill, border, etc.

#### Syntax
```js
rangeObject.clear(applyTo);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|applyTo|string|Optional. Determines the type of clear action. Possible values are: `All` Default-option,`Formats` ,`Contents` |

#### Returns
void

#### Examples

Below example clears format and contents of the range. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.clear();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### delete(shift: string)
Deletes the cells associated with the range.

#### Syntax
```js
rangeObject.delete(shift);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|shift|string|Specifies which way to shift the cells.  Possible values are: Up, Left|

#### Returns
void

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getBoundingRect(anotherRange: Range or string)
Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".

#### Syntax
```js
rangeObject.getBoundingRect(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|anotherRange|Range or string|The range object or address or range name.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D4:G6";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var range = range.getBoundingRect("G4:H8");
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // Prints Sheet1!D4:H8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getCell(row: number, column: number)
Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.

#### Syntax
```js
rangeObject.getCell(row, column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|row|number|Row number of the cell to be retrieved. Zero-indexed.|
|column|number|Column number of the cell to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var cell = range.cell(0,0);
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getColumn(column: number)
Gets a column contained in the range.

#### Syntax
```js
rangeObject.getColumn(column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|column|number|Column number of the range to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet19";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!B1:B8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getColumnsAfter(count: number)
Gets a certain number of columns to the right of the current Range object.

#### Syntax
```js
rangeObject.getColumnsAfter(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.|

#### Returns
[Range](range.md)

### getColumnsBefore(count: number)
Gets a certain number of columns to the left of the current Range object.

#### Syntax
```js
rangeObject.getColumnsBefore(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.|

#### Returns
[Range](range.md)

### getEntireColumn()
Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", it's `getEntireColumn` is a range that represents columns "B:E").

#### Syntax
```js
rangeObject.getEntireColumn();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

Note: the grid properties of the Range (values, numberFormat, formulas) contains `null` since the Range in question is unbounded.

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeEC = range.getEntireColumn();
	rangeEC.load('address');
	return ctx.sync().then(function() {
		console.log(rangeEC.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getEntireRow()
Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", it's `GetEntireRow` is a range that represents rows "4:11").

#### Syntax
```js
rangeObject.getEntireRow();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "D:F"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeER = range.getEntireRow();
	rangeER.load('address');
	return ctx.sync().then(function() {
		console.log(rangeER.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
The grid properties of the Range (values, numberFormat, formulas) contains `null` since the Range in question is unbounded.


### getIntersection(anotherRange: Range or string)
Gets the range object that represents the rectangular intersection of the given ranges.

#### Syntax
```js
rangeObject.getIntersection(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|anotherRange|Range or string|The range object or range address that will be used to determine the intersection of ranges.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!D4:F6
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getIntersectionOrNullObject(anotherRange: Range or string)
Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.

#### Syntax
```js
rangeObject.getIntersectionOrNullObject(anotherRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|anotherRange|Range or string|The range object or range address that will be used to determine the intersection of ranges.|

#### Returns
[Range](range.md)

### getLastCell()
Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".

#### Syntax
```js
rangeObject.getLastCell();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getLastColumn()
Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".

#### Syntax
```js
rangeObject.getLastColumn();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!F1:F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getLastRow()
Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".

#### Syntax
```js
rangeObject.getLastRow();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!A8:F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```



### getOffsetRange(rowOffset: number, columnOffset: number)
Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.

#### Syntax
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowOffset|number|The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.|
|columnOffset|number|The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.|

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D4:F6";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!H3:K5
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getResizedRange(deltaRows: number, deltaColumns: number)
Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.

#### Syntax
```js
rangeObject.getResizedRange(deltaRows, deltaColumns);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|deltaRows|number|The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.|
|deltaColumns|number|The number of columnsby which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.|

#### Returns
[Range](range.md)

### getRow(row: number)
Gets a row contained in the range.

#### Syntax
```js
rangeObject.getRow(row);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|row|number|Row number of the range to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!A2:F2
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getRowsAbove(count: number)
Gets a certain number of rows above the current Range object.

#### Syntax
```js
rangeObject.getRowsAbove(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.|

#### Returns
[Range](range.md)

### getRowsBelow(count: number)
Gets a certain number of rows below the current Range object.

#### Syntax
```js
rangeObject.getRowsBelow(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|Optional. The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.|

#### Returns
[Range](range.md)

### getUsedRange(valuesOnly: bool)
Returns the used range of the given range object. If there are no used cells within the range, this function will throw an ItemNotFound error.

#### Syntax
```js
rangeObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Considers only cells with values as used cells.|

#### Returns
[Range](range.md)

#### Examples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeUR = range.getUsedRange();
	rangeUR.load('address');
	return ctx.sync().then(function() {
		console.log(rangeUR.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getUsedRangeOrNullObject(valuesOnly: bool)
Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.

#### Syntax
```js
rangeObject.getUsedRangeOrNullObject(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Considers only cells with values as used cells.|

#### Returns
[Range](range.md)

### getVisibleView()
Represents the visible rows of the current range.

#### Syntax
```js
rangeObject.getVisibleView();
```

#### Parameters
None

#### Returns
[RangeView](rangeview.md)

### insert(shift: string)
Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.

#### Syntax
```js
rangeObject.insert(shift);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|shift|string|Specifies which way to shift the cells.  Possible values are: Down, Right|

#### Returns
[Range](range.md)

#### Examples

```js
	
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.insert();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### merge(across: bool)
Merge the range cells into one region in the worksheet.

#### Syntax
```js
rangeObject.merge(across);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|across|bool|Optional. Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.|

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:C3";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.merge(true);
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```



#### Examples
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:C3";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.unmerge();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### select()
Selects the specified range in the Excel UI.

#### Syntax
```js
rangeObject.select();
```

#### Parameters
None

#### Returns
void

#### Examples

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.select();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### unmerge()
Unmerge the range cells into separate cells.

#### Syntax
```js
rangeObject.unmerge();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:C3";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.unmerge();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Property access examples

Below example uses range address to get the range object.

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8"; 
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Below example uses a named-range to get the range object.

```js

Excel.run(function (ctx) { 
	var rangeName = 'MyRange';
	var range = ctx.workbook.names.getItem(rangeName).range;
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

The example below sets number-format, values and formulas on a grid that contains 2x3 grid.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:G7";
	var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
	var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
	var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.numberFormat = numberFormat;
	range.values = values;
	range.formulas= formulas;
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Get the worksheet containing the range. 

```js
/* This might be broken still - it was broken before because it 
	it was missing 'var', but might still be wrong because of
	getting information without loading properly. */
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	var range = namedItem.range;
	var rangeWorksheet = range.worksheet;
	rangeWorksheet.load('name');
	return ctx.sync().then(function() {
			console.log(rangeWorksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

