# Table Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a table in a Word document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[alignment](enums.md)|string|Gets or sets the alignment of the table against the page column. The value can be 'left', 'centered' or 'right'. Possible values are: `Unknown` Unknown alignment.,`Left` Alignment to the left.,`Centered` Alignment to the center.,`Right` Alignment to the right.,`Justified` Fully justified alignment.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|headerRowCount|int|Gets and sets the number of header rows.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[horizontalAlignment](enums.md)|string|Gets and sets the horizontal alignment of every cell in the table. The value can be 'left', 'centered', 'right', or 'justified'. Possible values are: `Unknown` Unknown alignment.,`Left` Alignment to the left.,`Centered` Alignment to the center.,`Right` Alignment to the right.,`Justified` Fully justified alignment.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|isUniform|bool|Indicates whether all of the table rows are uniform. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|nestingLevel|int|Gets the nesting level of the table. Top-level tables have level 1. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|rowCount|int|Gets the number of rows in the table. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|shadingColor|string|Gets and sets the shading color.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|style|string|Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|styleBandedColumns|bool|Gets and sets whether the table has banded columns.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|styleBandedRows|bool|Gets and sets whether the table has banded rows.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[styleBuiltIn](enums.md)|string|Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property. Possible values are: Other, Normal, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9, Toc1, more...|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|styleFirstColumn|bool|Gets and sets whether the table has a first column with a special style.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|styleLastColumn|bool|Gets and sets whether the table has a last column with a special style.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|styleTotalRow|bool|Gets and sets whether the table has a total (last) row with a special style.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|values|string|Gets and sets the text values in the table, as a 2D Javascript array.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[verticalAlignment](enums.md)|string|Gets and sets the vertical alignment of every cell in the table. The value can be 'top', 'center' or 'bottom'. Possible values are: Mixed, Top, Center, Bottom.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|width|float|Gets and sets the width of the table in points.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|font|[Font](font.md)|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentBody|[Body](body.md)|Gets the parent body of the table. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the table. Throws if there isn't a parent content control. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentContentControlOrNullObject|[ContentControl](contentcontrol.md)|Gets the content control that contains the table. Returns a null object if there isn't a parent content control. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentTable|[Table](table.md)|Gets the table that contains this table. Throws if it is not contained in a table. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains this table. Throws if it is not contained in a table cell. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentTableCellOrNullObject|[TableCell](tablecell.md)|Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|parentTableOrNullObject|[Table](table.md)|Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|rows|[TableRowCollection](tablerowcollection.md)|Gets all of the table rows. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Gets the child tables nested one level deeper. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[addColumns(insertLocation: string, columnCount: number, values: string[][])](#addcolumnsinsertlocation-string-columncount-number-values-string)|void|Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[addRows(insertLocation: string, rowCount: number, values: string[][])](#addrowsinsertlocation-string-rowcount-number-values-string)|[TableRowCollection](tablerowcollection.md)|Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[autoFitWindow()](#autofitwindow)|void|Autofits the table columns to the width of the window.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[clear()](#clear)|void|Clears the contents of the table.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[delete()](#delete)|void|Deletes the entire table.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[deleteColumns(columnIndex: number, columnCount: number)](#deletecolumnscolumnindex-number-columncount-number)|void|Deletes specific columns. This is applicable to uniform tables.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[deleteRows(rowIndex: number, rowCount: number)](#deleterowsrowindex-number-rowcount-number)|void|Deletes specific rows.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[distributeColumns()](#distributecolumns)|void|Distributes the column widths evenly. This is applicable to uniform tables.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getBorder(borderLocation: string)](#getborderborderlocation-string)|[TableBorder](tableborder.md)|Gets the border style for the specified border.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getCell(rowIndex: number, cellIndex: number)](#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Gets the table cell at a specified row and column. Throws if the specified table cell does not exist.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getCellOrNullObject(rowIndex: number, cellIndex: number)](#getcellornullobjectrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Gets the table cell at a specified row and column. Returns a null object if the specified table cell does not exist.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getCellPadding(cellPaddingLocation: CellPaddingLocation)](#getcellpaddingcellpaddinglocation-cellpaddinglocation)|[float?](float?.md)|Gets cell padding in points.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getNext()](#getnext)|[Table](table.md)|Gets the next table. Throws if this table is the last one.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getNextOrNullObject()](#getnextornullobject)|[Table](table.md)|Gets the next table. Returns a null object if this table is the last one.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getParagraphAfter()](#getparagraphafter)|[Paragraph](paragraph.md)|Gets the paragraph after the table. Throws if there isn't a paragraph after the table.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getParagraphAfterOrNullObject()](#getparagraphafterornullobject)|[Paragraph](paragraph.md)|Gets the paragraph after the table. Returns a null object if there isn't a paragraph after the table.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getParagraphBefore()](#getparagraphbefore)|[Paragraph](paragraph.md)|Gets the paragraph before the table. Throws if there isn't a paragraph before the table.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getParagraphBeforeOrNullObject()](#getparagraphbeforeornullobject)|[Paragraph](paragraph.md)|Gets the paragraph before the table. Returns a null object if there isn't a paragraph before the table.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getRange(rangeLocation: string)](#getrangerangelocation-string)|[Range](range.md)|Gets the range that contains this table, or the range at the start or end of the table.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Inserts a content control on the table.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])](#inserttablerowcount-number-columncount-number-insertlocation-string-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[RangeCollection](rangecollection.md)|Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[select(selectionMode: string)](#selectselectionmode-string)|void|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)](#setcellpaddingcellpaddinglocation-cellpaddinglocation-cellpadding-float)|void|Sets cell padding in points.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### addColumns(insertLocation: string, columnCount: number, values: string[][])
Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.

#### Syntax
```js
tableObject.addColumns(insertLocation, columnCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
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
|:---------------|:--------|:----------|
|insertLocation|string|Required. It can be 'Start' or 'End'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|rowCount|number|Required. Number of rows to add.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[TableRowCollection](tablerowcollection.md)

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
|:---------------|:--------|:----------|
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
|:---------------|:--------|:----------|
|rowIndex|number|Required. The first row to delete.|
|rowCount|number|Optional. Optional. The number of rows to delete. Default 1.|

#### Returns
void

### distributeColumns()
Distributes the column widths evenly. This is applicable to uniform tables.

#### Syntax
```js
tableObject.distributeColumns();
```

#### Parameters
None

#### Returns
void

### getBorder(borderLocation: string)
Gets the border style for the specified border.

#### Syntax
```js
tableObject.getBorder(borderLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|borderLocation|string|Required. The border location.  Possible values are: Top, Left, Bottom, Right, InsideHorizontal, InsideVertical, Inside, Outside, All|

#### Returns
[TableBorder](tableborder.md)

### getCell(rowIndex: number, cellIndex: number)
Gets the table cell at a specified row and column. Throws if the specified table cell does not exist.

#### Syntax
```js
tableObject.getCell(rowIndex, cellIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowIndex|number|Required. The index of the row.|
|cellIndex|number|Required. The index of the cell in the row.|

#### Returns
[TableCell](tablecell.md)

### getCellOrNullObject(rowIndex: number, cellIndex: number)
Gets the table cell at a specified row and column. Returns a null object if the specified table cell does not exist.

#### Syntax
```js
tableObject.getCellOrNullObject(rowIndex, cellIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowIndex|number|Required. The index of the row.|
|cellIndex|number|Required. The index of the cell in the row.|

#### Returns
[TableCell](tablecell.md)

### getCellPadding(cellPaddingLocation: CellPaddingLocation)
Gets cell padding in points.

#### Syntax
```js
tableObject.getCellPadding(cellPaddingLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cellPaddingLocation|CellPaddingLocation|Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.|

#### Returns
[float?](float?.md)

### getNext()
Gets the next table. Throws if this table is the last one.

#### Syntax
```js
tableObject.getNext();
```

#### Parameters
None

#### Returns
[Table](table.md)

### getNextOrNullObject()
Gets the next table. Returns a null object if this table is the last one.

#### Syntax
```js
tableObject.getNextOrNullObject();
```

#### Parameters
None

#### Returns
[Table](table.md)

### getParagraphAfter()
Gets the paragraph after the table. Throws if there isn't a paragraph after the table.

#### Syntax
```js
tableObject.getParagraphAfter();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

### getParagraphAfterOrNullObject()
Gets the paragraph after the table. Returns a null object if there isn't a paragraph after the table.

#### Syntax
```js
tableObject.getParagraphAfterOrNullObject();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

### getParagraphBefore()
Gets the paragraph before the table. Throws if there isn't a paragraph before the table.

#### Syntax
```js
tableObject.getParagraphBefore();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

### getParagraphBeforeOrNullObject()
Gets the paragraph before the table. Returns a null object if there isn't a paragraph before the table.

#### Syntax
```js
tableObject.getParagraphBeforeOrNullObject();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

### getRange(rangeLocation: string)
Gets the range that contains this table, or the range at the start or end of the table.

#### Syntax
```js
tableObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|string|Optional. Optional. The range location can be 'Whole', 'Start', 'End' or 'After'.  Possible values are: Whole, Start, End, Before, After, Content|

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
|:---------------|:--------|:----------|
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
|:---------------|:--------|:----------|
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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.

#### Syntax
```js
tableObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[RangeCollection](rangecollection.md)

### select(selectionMode: string)
Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.

#### Syntax
```js
tableObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|string|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.  Possible values are: Select, Start, End|

#### Returns
void

### setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)
Sets cell padding in points.

#### Syntax
```js
tableObject.setCellPadding(cellPaddingLocation, cellPadding);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|cellPaddingLocation|CellPaddingLocation|Required. The cell padding location can be 'Top', 'Left', 'Bottom' or 'Right'.|
|cellPadding|float|Required. The cell padding.|

#### Returns
void
