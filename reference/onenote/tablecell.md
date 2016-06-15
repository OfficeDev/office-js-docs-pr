# TableCell Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a cell in a OneNote table.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|:-------|
|cellIndex|int|Gets the index of the cell in its row. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-cellIndex)|
|id|string|Gets the ID of the cell. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-id)|
|rowIndex|int|Gets the index of the cell's row in the table. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-rowIndex)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|:-------|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of Paragraph objects in the TableCell. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-paragraphs)|
|parentRow|[TableRow](tablerow.md)|Gets the parent row of the cell. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-parentRow)|
|parentTable|[Table](table.md)|Gets the parent table of the cell. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-parentTable)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Adds the specified HTML to the bottom of the TableCell.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendHtml)|
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
				cell.appendHtml("Hello");
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

