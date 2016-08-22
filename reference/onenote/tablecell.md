# TableCell Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents a cell in a OneNote table.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|cellIndex|int|Gets the index of the cell in its row. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-cellIndex)|
|id|string|Gets the ID of the cell. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-id)|
|rowIndex|int|Gets the index of the cell's row in the table. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-rowIndex)|
|shadingColor|string|Gets and sets the shading color of the cell|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-shadingColor)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of Paragraph objects in the TableCell. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-paragraphs)|
|parentRow|[TableRow](tablerow.md)|Gets the parent row of the cell. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-parentRow)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Adds the specified HTML to the bottom of the TableCell.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|Adds the specified image to table cell.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|Adds the specified text to table cell.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|Adds a table with the specified number of rows and columns to table cell.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendTable)|
|[clear()](#clear)|void|Clears the contents of the cell.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-clear)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-load)|

## Method Details


### appendHtml(html: string)
Adds the specified HTML to the bottom of the TableCell.

#### Syntax
```js
tableCellObject.appendHtml(html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|The HTML string to append. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.|

#### Returns
void

#### Examples
```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, get a table cell at row one and column two and add "Hello".
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
				cell.appendHtml("<p>Hello</p>");
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});


### appendImage(base64EncodedImage: string, width: double, height: double)
Adds the specified image to table cell.

#### Syntax
```js
tableCellObject.appendImage(base64EncodedImage, width, height);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|HTML string to append.|
|width|double|Optional. Width in the unit of Points. The default value is null and image width will be respected.|
|height|double|Optional. Height in the unit of Points. The default value is null and image height will be respected.|

#### Returns
[Image](image.md)

### appendRichText(paragraphText: string)
Adds the specified text to table cell.

#### Syntax
```js
tableCellObject.appendRichText(paragraphText);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|HTML string to append.|

#### Returns
[RichText](richtext.md)

#### Examples
```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	var appendedRichText = null;
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, get a table cell at row one and column two and add "Hello".
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
				appendedRichText = cell.appendRichText("Hello");
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

### appendTable(rowCount: number, columnCount: number, values: string[][])
Adds a table with the specified number of rows and columns to table cell.

#### Syntax
```js
tableCellObject.appendTable(rowCount, columnCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowCount|number|Required. The number of rows in the table.|
|columnCount|number|Required. The number of columns in the table.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[Table](table.md)

### clear()
Clears the contents of the cell.

#### Syntax
```js
tableCellObject.clear();
```

#### Parameters
None

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
### Property access examples
**id, cellIndex, rowIndex**
```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, get a table cell at row one and column two.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
				
				// Queue a command to load the table cell.
				ctx.load(cell);
				ctx.sync().then(function() {
					console.log("Cell Id: " + cell.id);
					console.log("Cell Index: " + cell.cellIndex);
					console.log("Cell's Row Index: " + cell.rowIndex);
				});
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

**parentTable, cells**
```js
ParentTable, ParentRow, Paragraphs
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, get a table cell at row one and column two.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
				
				// Queue a command to load parentTable, parentRow and paragraphs of the table cell.
				ctx.load(cell, "parentTable, parentRow, paragraphs");
				
				ctx.sync().then(function() {
					console.log("Parent Table Id: " + cell.parentTable.id);
					console.log("Parent Row Id: " + cell.parentRow.id);
					var paragraphs = cell.paragraphs;
					
					for (var i = 0; i < paragraphs.items.length; i++) {
						console.log("Paragraph Id: " + paragraphs.items[i].id);
					}
				});
			}
		}
		return ctx.sync();
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

