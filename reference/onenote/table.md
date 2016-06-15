# Table Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents a table in a OneNote page.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|columnCount|int|Gets the number of columns in the table. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-columnCount)|
|id|string|Gets the ID of the table. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-id)|
|rowCount|int|Gets the number of rows in the table. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rowCount)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the Table object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-paragraph)|
|rows|[TableRowCollection](tablerowcollection.md)|Gets all of the table rows. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rows)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[appendColumn(values: string[])](#appendcolumnvalues-string)|void|Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise the column is empty.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendColumn)|
|[appendRow(values: string[])](#appendrowvalues-string)|[TableRow](tablerow.md)|Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise the row is empty.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendRow)|
|[deleteColumns(columnIndex: number, columnCount: number)](#deletecolumnscolumnindex-number-columncount-number)|void|Deletes a contiguous run of columns.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-deleteColumns)|
|[deleteRows(rowIndex: number, rowCount: number)](#deleterowsrowindex-number-rowcount-number)|void|Deletes a contiguous run of rows.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-deleteRows)|
|[getCell(rowIndex: number, cellIndex: number)](#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Gets the table cell at a specified row and column.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-getCell)|
|[hideBorder()](#hideborder)|void|Hides the table's border|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-hideBorder)|
|[insertColumn(index: number, values: string[])](#insertcolumnindex-number-values-string)|void|Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise the column is empty.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertColumn)|
|[insertRow(index: number, values: string[])](#insertrowindex-number-values-string)|[TableRow](tablerow.md)|Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise the row is empty.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertRow)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-load)|
|[showBorder()](#showborder)|void|Make's the table's border visible|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-showBorder)|

## Method Details


### appendColumn(values: string[])
Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise the column is empty.

#### Syntax
```js
tableObject.appendColumn(values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|string[]|Optional. Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.|

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
		
		// for each table, append a column.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				table.appendColumn(["cell0", "cell1"]);
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


### appendRow(values: string[])
Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise the row is empty.

#### Syntax
```js
tableObject.appendRow(values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|string[]|Optional. Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.|

#### Returns
[TableRow](tablerow.md)

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
		
		// for each table, append a column.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var row = table.appendRow(["cell0", "cell1"]);
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


### deleteColumns(columnIndex: number, columnCount: number)
Deletes a contiguous run of columns.

#### Syntax
```js
tableObject.deleteColumns(columnIndex, columnCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|columnIndex|number|The first column to delete.|
|columnCount|number|Optional. Optional. The number of columns to delete. Default 1.|

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
		
		// for each table, delete columns.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				table.deleteColumns(2 /*Index of column to delete*/, 3 /*Number of columns to delete*/);
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


### deleteRows(rowIndex: number, rowCount: number)
Deletes a contiguous run of rows.

#### Syntax
```js
tableObject.deleteRows(rowIndex, rowCount);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowIndex|number|The first row to delete.|
|rowCount|number|Optional. Optional. The number of rows to delete. Default 1.|

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
		
		// for each table, delete rows.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				table.deleteRows(2 /*Index of row to delete*/, 3 /*Number of rows to delete*/);
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


### getCell(rowIndex: number, cellIndex: number)
Gets the table cell at a specified row and column.

#### Syntax
```js
tableObject.getCell(rowIndex, cellIndex);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowIndex|number|The index of the row.|
|cellIndex|number|The index of the cell in the row.|

#### Returns
[TableCell](tablecell.md)

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
		
		// for each table, get a cell in the second row and third column.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var cell = table.getCell(2 /*Row Index*/, 3 /*Column Index*/);
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


### hideBorder()
Hides the table's border

#### Syntax
```js
tableObject.hideBorder();
```

#### Parameters
None

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
		
		// for each table, hide border.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				table.hideBorder();
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


### insertColumn(index: number, values: string[])
Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise the column is empty.

#### Syntax
```js
tableObject.insertColumn(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index where the column will be inserted in the table.|
|values|string[]|Optional. Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.|

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
		
		// for each table, insert a column at index two.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				table.insertColumn(2, ["cell0", "cell1"]);
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


### insertRow(index: number, values: string[])
Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise the row is empty.

#### Syntax
```js
tableObject.insertRow(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index where the row will be inserted in the table.|
|values|string[]|Optional. Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.|

#### Returns
[TableRow](tablerow.md)

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
		
		// for each table, insert a row at index two.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				var row = table.insertRow(2, ["cell0", "cell1"]);
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

### showBorder()
Make's the table's border visible

#### Syntax
```js
tableObject.showBorder();
```

#### Parameters
None

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
		
		// for each table, show border.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				table.showBorder();
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
### Property access examples
**columnCount, rowCount, id**
```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// For each table, log properties.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				ctx.load(table);
				return ctx.sync().then(function() {
					console.log("Table Id: " + table.id);
					console.log("Row Count: " + table.rowCount);
					console.log("Column Count: " + table.columnCount);
					return ctx.sync();
				});
			}
		}
	});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

** paragraph, rows **
```js
OneNote.run(function(ctx) {
	var app = ctx.application;
	var outline = app.getActiveOutline();
	
	// Queue a command to load outline.paragraphs and their types.
	ctx.load(outline, "paragraphs, paragraphs/type");
	
	// Run the queued commands, and return a promise to indicate task completion.
	return ctx.sync().then(function () {
		var paragraphs = outline.paragraphs;
		
		// for each table, log its paragraph id.
		for (var i = 0; i < paragraphs.items.length; i++) {
			var paragraph = paragraphs.items[i];
			if (paragraph.type == "Table") {
				var table = paragraph.table;
				ctx.load(table, "paragraph/id, rows/id");
				return ctx.sync().then(function() {
					console.log("Paragraph Id: " + table.paragraph.id);
					var rows = table.rows;
					
					// for each rows in the table, log row index and id.
					for (var i = 0; i < rows.items.length; i++) {
						console.log("Row " + i + " Id: " + rows.items[i].id);
					}
					return ctx.sync();
				});
			}
		}
	})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

