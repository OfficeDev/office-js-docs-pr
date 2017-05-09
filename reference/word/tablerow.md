# TableRow Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a row in a Word document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|cellCount|int|Gets the number of cells in the row. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[horizontalAlignment](enums.md)|string|Gets and sets the horizontal alignment of every cell in the row. The value can be 'left', 'centered', 'right', or 'justified'. Possible values are: `Unknown` Unknown alignment.,`Left` Alignment to the left.,`Centered` Alignment to the center.,`Right` Alignment to the right.,`Justified` Fully justified alignment.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|isHeader|bool|Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|preferredHeight|float|Gets and sets the preferred height of the row in points.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|rowIndex|int|Gets the index of the row in its parent table. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|shadingColor|string|Gets and sets the shading color.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|values|string|Gets and sets the text values in the row, as a 2D Javascript array.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[verticalAlignment](enums.md)|string|Gets and sets the vertical alignment of the cells in the row. The value can be 'top', 'center' or 'bottom'. Possible values are: Mixed, Top, Center, Bottom.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|cells|[TableCellCollection](tablecellcollection.md)|Gets cells. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|font|[Font](font.md)|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentTable|[Table](table.md)|Gets parent table. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clears the contents of the row.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[delete()](#delete)|void|Deletes the entire row.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getBorder(borderLocation: string)](#getborderborderlocation-string)|[TableBorder](tableborder.md)|Gets the border style of the cells in the row.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getCellPadding(cellPaddingLocation: CellPaddingLocation)](#getcellpaddingcellpaddinglocation-cellpaddinglocation)|[float?](float?.md)|Gets cell padding in points.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getNext()](#getnext)|[TableRow](tablerow.md)|Gets the next row. Throws if this row is the last one.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getNextOrNullObject()](#getnextornullobject)|[TableRow](tablerow.md)|Gets the next row. Returns a null object if this row is the last one.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[insertRows(insertLocation: string, rowCount: number, values: string[][])](#insertrowsinsertlocation-string-rowcount-number-values-string)|[TableRowCollection](tablerowcollection.md)|Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[merge()](#merge)|[TableCell](tablecell.md)|Merges the row into one cell.|[1.4](../requirement-sets/word-api-requirement-sets.md)|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[RangeCollection](rangecollection.md)|Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[select(selectionMode: string)](#selectselectionmode-string)|void|Selects the row and navigates the Word UI to it.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)](#setcellpaddingcellpaddinglocation-cellpaddinglocation-cellpadding-float)|void|Sets cell padding in points.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### clear()
Clears the contents of the row.

#### Syntax
```js
tableRowObject.clear();
```

#### Parameters
None

#### Returns
void

### delete()
Deletes the entire row.

#### Syntax
```js
tableRowObject.delete();
```

#### Parameters
None

#### Returns
void

### getBorder(borderLocation: string)
Gets the border style of the cells in the row.

#### Syntax
```js
tableRowObject.getBorder(borderLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|borderLocation|string|Required. The border location.  Possible values are: Top, Left, Bottom, Right, InsideHorizontal, InsideVertical, Inside, Outside, All|

#### Returns
[TableBorder](tableborder.md)

### getCellPadding(cellPaddingLocation: CellPaddingLocation)
Gets cell padding in points.

#### Syntax
```js
tableRowObject.getCellPadding(cellPaddingLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|cellPaddingLocation|CellPaddingLocation|Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.|

#### Returns
[float?](float?.md)

### getNext()
Gets the next row. Throws if this row is the last one.

#### Syntax
```js
tableRowObject.getNext();
```

#### Parameters
None

#### Returns
[TableRow](tablerow.md)

### getNextOrNullObject()
Gets the next row. Returns a null object if this row is the last one.

#### Syntax
```js
tableRowObject.getNextOrNullObject();
```

#### Parameters
None

#### Returns
[TableRow](tablerow.md)

### insertRows(insertLocation: string, rowCount: number, values: string[][])
Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.

#### Syntax
```js
tableRowObject.insertRows(insertLocation, rowCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|insertLocation|string|Required. Where the new rows should be inserted, relative to the current row. It can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|rowCount|number|Required. Number of rows to add|
|values|string[][]|Optional. Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.|

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
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### merge()
Merges the row into one cell.

#### Syntax
```js
tableRowObject.merge();
```

#### Parameters
None

#### Returns
[TableCell](tablecell.md)

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.

#### Syntax
```js
tableRowObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[RangeCollection](rangecollection.md)

### select(selectionMode: string)
Selects the row and navigates the Word UI to it.

#### Syntax
```js
tableRowObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|selectionMode|string|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.  Possible values are: Select, Start, End|

#### Returns
void

### setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)
Sets cell padding in points.

#### Syntax
```js
tableRowObject.setCellPadding(cellPaddingLocation, cellPadding);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|cellPaddingLocation|CellPaddingLocation|Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.|
|cellPadding|float|Required. The cell padding.|

#### Returns
void
