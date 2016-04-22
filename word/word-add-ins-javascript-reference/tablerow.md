# TableRow Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a row in a Word document.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|cellCount|int|Gets the number of cells in the row. Read-only.|
|isHeader|bool|Gets a value that indicates whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object. Read-only.|
|rowIndex|int|Gets the index of the row in its parent table. Read-only.|
|shadingColor|string|Gets and sets the shading color.|
|values|string|Gets and sets the text values in the row, as a 1D Javascript array.|
|verticalAlignment|string|Gets and sets the vertical alignment of the cells in the row. Possible values are: Mixed, Top, Center, Bottom.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|cellPaddingBottom|[float](float.md)|Gets and sets the default bottom cell padding for the row in points.|
|cellPaddingLeft|[float](float.md)|Gets and sets the default left cell padding for the row in points.|
|cellPaddingRight|[float](float.md)|Gets and sets the default right cell padding for the row in points.|
|cellPaddingTop|[float](float.md)|Gets and sets the default top cell padding for the row in points.|
|cells|[TableCellCollection](tablecellcollection.md)|Gets cells. Read-only.|
|font|[Font](font.md)|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|
|next|[TableRow](tablerow.md)|Gets the next row. Read-only.|
|parentTable|[Table](table.md)|Gets parent table. Read-only.|
|preferredHeight|[float](float.md)|Gets and sets the preferred height of the row in points.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the row.|
|[delete()](#delete)|void|Deletes the entire row.|
|[getBorderStyle(borderLocation: string)](#getborderstyleborderlocation-string)|[TableBorderStyle](tableborderstyle.md)|Gets the border style of the cells in the row.|
|[insertRows(insertLocation: string, rowCount: number, values: string[][])](#insertrowsinsertlocation-string-rowcount-number-values-string)|void|Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[merge()](#merge)|[TableCell](tablecell.md)|Merges the row into one cell.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.|
|[select(selectionMode: string)](#selectselectionmode-string)|void|Selects the row and navigates the Word UI to it.|

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

### getBorderStyle(borderLocation: string)
Gets the border style of the cells in the row.

#### Syntax
```js
tableRowObject.getBorderStyle(borderLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|borderLocation|string|Required. The border location.  Possible values are: Top, Left, Bottom, Right, InsideHorizontal, InsideVertical, Inside, Outside, All|

#### Returns
[TableBorderStyle](tableborderstyle.md)

### insertRows(insertLocation: string, rowCount: number, values: string[][])
Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.

#### Syntax
```js
tableRowObject.insertRows(insertLocation, rowCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|Where the new rows should be inserted, relative to the current row. It can be 'Before' or 'After'. Required. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|rowCount|number|Required. Number of rows to add|
|values|string[][]|Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row. Optional.|

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
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: string)
Selects the row and navigates the Word UI to it.

#### Syntax
```js
tableRowObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|string|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.  Possible values are: Select, Start, End|

#### Returns
void
