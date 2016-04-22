# TableCell Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a table cell in a Word document.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|cellIndex|int|Gets the index of the cell in its row. Read-only.|
|rowIndex|int|Gets the index of the cell's row in the table. Read-only.|
|shadingColor|string|Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.|
|value|string|Gets and sets the text of the cell.|
|verticalAlignment|string|Gets and sets the vertical alignment of the cell. Possible values are: Mixed, Top, Center, Bottom.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|body|[Body](body.md)|Gets the body object of the cell. Read-only.|
|cellPaddingBottom|[float](float.md)|Gets and sets the bottom padding of the cell in points.|
|cellPaddingLeft|[float](float.md)|Gets and sets the left padding of the cell in points.|
|cellPaddingRight|[float](float.md)|Gets and sets the right padding of the cell in points.|
|cellPaddingTop|[float](float.md)|Gets and sets the top padding of the cell in points.|
|columnWidth|[float](float.md)|Gets and sets the width of the cell's column in points. This is applicable to uniform tables.|
|next|[TableCell](tablecell.md)|Gets the next cell. Read-only.|
|parentRow|[TableRow](tablerow.md)|Gets the parent row of the cell. Read-only.|
|parentTable|[Table](table.md)|Gets the parent table of the cell. Read-only.|
|width|[float](float.md)|Gets the width of the cell in points. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[deleteColumn()](#deletecolumn)|void|Deletes the column containing this cell. This is applicable to uniform tables.|
|[deleteRow()](#deleterow)|void|Deletes the row containing this cell.|
|[getBorderStyle(borderLocation: string)](#getborderstyleborderlocation-string)|[TableBorderStyle](tableborderstyle.md)|Gets the border style for the specified border.|
|[insertColumns(insertLocation: string, columnCount: number, values: string[][])](#insertcolumnsinsertlocation-string-columncount-number-values-string)|void|Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|
|[insertRows(insertLocation: string, rowCount: number, values: string[][])](#insertrowsinsertlocation-string-rowcount-number-values-string)|void|Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[split(rowCount: number, columnCount: number)](#splitrowcount-number-columncount-number)|void|Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.|

## Method Details


### deleteColumn()
Deletes the column containing this cell. This is applicable to uniform tables.

#### Syntax
```js
tableCellObject.deleteColumn();
```

#### Parameters
None

#### Returns
void

### deleteRow()
Deletes the row containing this cell.

#### Syntax
```js
tableCellObject.deleteRow();
```

#### Parameters
None

#### Returns
void

### getBorderStyle(borderLocation: string)
Gets the border style for the specified border.

#### Syntax
```js
tableCellObject.getBorderStyle(borderLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|borderLocation|string|Required. The border location.  Possible values are: Top, Left, Bottom, Right, InsideHorizontal, InsideVertical, Inside, Outside, All|

#### Returns
[TableBorderStyle](tableborderstyle.md)

### insertColumns(insertLocation: string, columnCount: number, values: string[][])
Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableCellObject.insertColumns(insertLocation, columnCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|Required. It can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|columnCount|number|Required. Number of columns to add|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
void

### insertRows(insertLocation: string, rowCount: number, values: string[][])
Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableCellObject.insertRows(insertLocation, rowCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|Required. It can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|rowCount|number|Required. Number of rows to add.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
void

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

### split(rowCount: number, columnCount: number)
Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableCellObject.split(rowCount, columnCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowCount|number|Required. The number of rows to split into. Must be a divisor of the number of underlying rows.|
|columnCount|number|Required. The number of columns to split into.|

#### Returns
void
