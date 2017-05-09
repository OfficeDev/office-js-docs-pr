# TableCell Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a table cell in a Word document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|cellIndex|int|Gets the index of the cell in its row. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|columnWidth|float|Gets and sets the width of the cell's column in points. This is applicable to uniform tables.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[horizontalAlignment](enums.md)|string|Gets and sets the horizontal alignment of the cell. The value can be 'left', 'centered', 'right', or 'justified'. Possible values are: `Unknown` Unknown alignment.,`Left` Alignment to the left.,`Centered` Alignment to the center.,`Right` Alignment to the right.,`Justified` Fully justified alignment.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|rowIndex|int|Gets the index of the cell's row in the table. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|shadingColor|string|Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|value|string|Gets and sets the text of the cell.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[verticalAlignment](enums.md)|string|Gets and sets the vertical alignment of the cell. The value can be 'top', 'center' or 'bottom'. Possible values are: Mixed, Top, Center, Bottom.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|width|float|Gets the width of the cell in points. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|body|[Body](body.md)|Gets the body object of the cell. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentRow|[TableRow](tablerow.md)|Gets the parent row of the cell. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentTable|[Table](table.md)|Gets the parent table of the cell. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[deleteColumn()](#deletecolumn)|void|Deletes the column containing this cell. This is applicable to uniform tables.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[deleteRow()](#deleterow)|void|Deletes the row containing this cell.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getBorder(borderLocation: string)](#getborderborderlocation-string)|[TableBorder](tableborder.md)|Gets the border style for the specified border.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getCellPadding(cellPaddingLocation: CellPaddingLocation)](#getcellpaddingcellpaddinglocation-cellpaddinglocation)|[float?](float?.md)|Gets cell padding in points.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getNext()](#getnext)|[TableCell](tablecell.md)|Gets the next cell. Throws if this cell is the last one.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getNextOrNullObject()](#getnextornullobject)|[TableCell](tablecell.md)|Gets the next cell. Returns a null object if this cell is the last one.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[insertColumns(insertLocation: string, columnCount: number, values: string[][])](#insertcolumnsinsertlocation-string-columncount-number-values-string)|void|Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[insertRows(insertLocation: string, rowCount: number, values: string[][])](#insertrowsinsertlocation-string-rowcount-number-values-string)|[TableRowCollection](tablerowcollection.md)|Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)](#setcellpaddingcellpaddinglocation-cellpaddinglocation-cellpadding-float)|void|Sets cell padding in points.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

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

### getBorder(borderLocation: string)
Gets the border style for the specified border.

#### Syntax
```js
tableCellObject.getBorder(borderLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|borderLocation|string|Required. The border location.  Possible values are: Top, Left, Bottom, Right, InsideHorizontal, InsideVertical, Inside, Outside, All|

#### Returns
[TableBorder](tableborder.md)

### getCellPadding(cellPaddingLocation: CellPaddingLocation)
Gets cell padding in points.

#### Syntax
```js
tableCellObject.getCellPadding(cellPaddingLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cellPaddingLocation|CellPaddingLocation|Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.|

#### Returns
[float?](float?.md)

### getNext()
Gets the next cell. Throws if this cell is the last one.

#### Syntax
```js
tableCellObject.getNext();
```

#### Parameters
None

#### Returns
[TableCell](tablecell.md)

### getNextOrNullObject()
Gets the next cell. Returns a null object if this cell is the last one.

#### Syntax
```js
tableCellObject.getNextOrNullObject();
```

#### Parameters
None

#### Returns
[TableCell](tablecell.md)

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
[TableRowCollection](tablerowcollection.md)

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

### setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)
Sets cell padding in points.

#### Syntax
```js
tableCellObject.setCellPadding(cellPaddingLocation, cellPadding);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cellPaddingLocation|CellPaddingLocation|Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.|
|cellPadding|float|Required. The cell padding.|

#### Returns
void
