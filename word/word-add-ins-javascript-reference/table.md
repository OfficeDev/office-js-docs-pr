# Table Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a table in a Word document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|headerRowCount|int|Gets and sets the number of header rows.|WordApi1.3||
|isUniform|bool|Indicates whether all of the table rows are uniform. Read-only.|WordApi1.3||
|nestingLevel|int|Gets the nesting level of the table. Top-level tables have level 1. Read-only.|WordApi1.3||
|rowCount|int|Gets the number of rows in the table. Read-only.|WordApi1.3||
|shadingColor|string|Gets and sets the shading color.|WordApi1.3||
|style|string|Gets and sets the name of the table style.|WordApi1.3||
|styleBandedColumns|bool|Gets and sets whether the table has banded columns.|WordApi1.3||
|styleBandedRows|bool|Gets and sets whether the table has banded rows.|WordApi1.3||
|styleFirstColumn|bool|Gets and sets whether the table has a first column with a special style.|WordApi1.3||
|styleLastColumn|bool|Gets and sets whether the table has a last column with a special style.|WordApi1.3||
|styleTotalRow|bool|Gets and sets whether the table has a total (last) row with a special style.|WordApi1.3||
|values|string|Gets and sets the text values in the table, as a 2D Javascript array.|WordApi1.3||
|verticalAlignment|string|Gets and sets the vertical alignment of every cell in the table. Possible values are: Mixed, Top, Center, Bottom.|WordApi1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|cellPaddingBottom|[float](float.md)|Gets and sets the default bottom cell padding in points.|WordApi1.3||
|cellPaddingLeft|[float](float.md)|Gets and sets the default left cell padding in points.|WordApi1.3||
|cellPaddingRight|[float](float.md)|Gets and sets the default right cell padding in points.|WordApi1.3||
|cellPaddingTop|[float](float.md)|Gets and sets the default top cell padding in points.|WordApi1.3||
|font|[Font](font.md)|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|WordApi1.3||
|height|[float](float.md)|Gets the height of the table in points. Read-only.|WordApi1.3||
|next|[Table](table.md)|Gets the next table. Read-only.|WordApi1.3||
|paragraphAfter|[Paragraph](paragraph.md)|Gets the paragraph after the table. Read-only.|WordApi1.3||
|paragraphBefore|[Paragraph](paragraph.md)|Gets the paragraph before the table. Read-only.|WordApi1.3||
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the table. Read-only.|WordApi1.3||
|parentTable|[Table](table.md)|Gets the table that contains this table. Returns null if it is not contained in a table. Read-only.|WordApi1.3||
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains this table. Returns null if it is not contained in a table cell. Read-only.|WordApi1.3||
|rows|[TableRowCollection](tablerowcollection.md)|Gets all of the table rows. Read-only.|WordApi1.3||
|tables|[TableCollection](tablecollection.md)|Gets the child tables nested one level deeper. Read-only.|WordApi1.3||
|width|[float](float.md)|Gets and sets the width of the table in points.|WordApi1.3||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[addColumns(insertLocation: string, columnCount: number, values: string[][])](#addcolumnsinsertlocation-string-columncount-number-values-string)|void|Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|WordApi1.3|
|[addRows(insertLocation: string, rowCount: number, values: string[][])](#addrowsinsertlocation-string-rowcount-number-values-string)|void|Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.|WordApi1.3|
|[autoFitContents()](#autofitcontents)|void|Autofits the table columns to the width of their contents.|WordApi1.3|
|[autoFitWindow()](#autofitwindow)|void|Autofits the table columns to the width of the window.|WordApi1.3|
|[clear()](#clear)|void|Clears the contents of the table.|WordApi1.3|
|[delete()](#delete)|void|Deletes the entire table.|WordApi1.3|
|[deleteColumns(columnIndex: number, columnCount: number)](#deletecolumnscolumnindex-number-columncount-number)|void|Deletes specific columns. This is applicable to uniform tables.|WordApi1.3|
|[deleteRows(rowIndex: number, rowCount: number)](#deleterowsrowindex-number-rowcount-number)|void|Deletes specific rows.|WordApi1.3|
|[distributeColumns()](#distributecolumns)|void|Distributes the column widths evenly.|WordApi1.3|
|[distributeRows()](#distributerows)|void|Distributes the row heights evenly.|WordApi1.3|
|[getBorderStyle(borderLocation: string)](#getborderstyleborderlocation-string)|[TableBorderStyle](tableborderstyle.md)|Gets the border style for the specified border.|WordApi1.3|
|[getCell(rowIndex: number, cellIndex: number)](#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Gets the table cell at a specified row and column.|WordApi1.3|
|[getRange(rangeLocation: string)](#getrangerangelocation-string)|[Range](range.md)|Gets the range that contains this table, or the range at the start or end of the table.|WordApi1.3|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Inserts a content control on the table.|WordApi1.3|
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|WordApi1.3|
|[insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])](#inserttablerowcount-number-columncount-number-insertlocation-string-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|WordApi1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|WordApi1.1|
|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](#mergecellstoprow-number-firstcell-number-bottomrow-number-lastcell-number)|[TableCell](tablecell.md)|Merges the cells bounded inclusively by a first and last cell.|WordApi1.1|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.|WordApi1.3|
|[select(selectionMode: string)](#selectselectionmode-string)|void|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|WordApi1.3|

## Method Details


### addColumns(insertLocation: string, columnCount: number, values: string[][])
Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableObject.addColumns(insertLocation, columnCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|insertLocation|string|Required. It can be 'Start' or 'End', corresponding to the appropriate side of the table. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|columnCount|number|Required. Number of columns to add.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
void

### addRows(insertLocation: string, rowCount: number, values: string[][])
Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableObject.addRows(insertLocation, rowCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|insertLocation|string|Required. It can be 'Start' or 'End'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|rowCount|number|Required. Number of rows to add.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
void

### autoFitContents()
Autofits the table columns to the width of their contents.

#### Syntax
```js
tableObject.autoFitContents();
```

#### Parameters
None

#### Returns
void

### autoFitWindow()
Autofits the table columns to the width of the window.

#### Syntax
```js
tableObject.autoFitWindow();
```

#### Parameters
None

#### Returns
void

### clear()
Clears the contents of the table.

#### Syntax
```js
tableObject.clear();
```

#### Parameters
None

#### Returns
void

### delete()
Deletes the entire table.

#### Syntax
```js
tableObject.delete();
```

#### Parameters
None

#### Returns
void

### deleteColumns(columnIndex: number, columnCount: number)
Deletes specific columns. This is applicable to uniform tables.

#### Syntax
```js
tableObject.deleteColumns(columnIndex, columnCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|columnIndex|number|Required. The first column to delete.|
|columnCount|number|Optional. Optional. The number of columns to delete. Default 1.|

#### Returns
void

### deleteRows(rowIndex: number, rowCount: number)
Deletes specific rows.

#### Syntax
```js
tableObject.deleteRows(rowIndex, rowCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rowIndex|number|Required. The first row to delete.|
|rowCount|number|Optional. Optional. The number of rows to delete. Default 1.|

#### Returns
void

### distributeColumns()
Distributes the column widths evenly.

#### Syntax
```js
tableObject.distributeColumns();
```

#### Parameters
None

#### Returns
void

### distributeRows()
Distributes the row heights evenly.

#### Syntax
```js
tableObject.distributeRows();
```

#### Parameters
None

#### Returns
void

### getBorderStyle(borderLocation: string)
Gets the border style for the specified border.

#### Syntax
```js
tableObject.getBorderStyle(borderLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|borderLocation|string|Required. The border location.  Possible values are: Top, Left, Bottom, Right, InsideHorizontal, InsideVertical, Inside, Outside, All|

#### Returns
[TableBorderStyle](tableborderstyle.md)

### getCell(rowIndex: number, cellIndex: number)
Gets the table cell at a specified row and column.

#### Syntax
```js
tableObject.getCell(rowIndex, cellIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rowIndex|number|Required. The index of the row.|
|cellIndex|number|Required. The index of the cell in the row.|

#### Returns
[TableCell](tablecell.md)

### getRange(rangeLocation: string)
Gets the range that contains this table, or the range at the start or end of the table.

#### Syntax
```js
tableObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rangeLocation|string|Optional. The range location can be 'Whole', 'Start' or 'End'.  Possible values are: Whole, Start, End|

#### Returns
[Range](range.md)

### insertContentControl()
Inserts a content control on the table.

#### Syntax
```js
tableObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

### insertParagraph(paragraphText: string, insertLocation: string)
Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
tableObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Paragraph](paragraph.md)

### insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])
Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
tableObject.insertTable(rowCount, columnCount, insertLocation, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rowCount|number|Required. The number of rows in the table.|
|columnCount|number|Required. The number of columns in the table.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[Table](table.md)

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

### mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)
Merges the cells bounded inclusively by a first and last cell.

#### Syntax
```js
tableObject.mergeCells(topRow, firstCell, bottomRow, lastCell);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|topRow|number|Required. The row of the first cell|
|firstCell|number|Required. The index of the first cell in its row|
|bottomRow|number|Required. The row of the last cell|
|lastCell|number|Required. The index of the last cell in its row|

#### Returns
[TableCell](tablecell.md)

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.

#### Syntax
```js
tableObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: string)
Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.

#### Syntax
```js
tableObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|selectionMode|string|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.  Possible values are: Select, Start, End|

#### Returns
void
