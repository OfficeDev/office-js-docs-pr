# BindingCollection object (JavaScript API for Excel)

Represents the collection of all the binding objects that are part of the workbook.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of bindings in the collection. Read-only.|
|items|[Binding[]](binding.md)|A collection of binding objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|Gets a binding object by ID.|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|Gets a binding object based on its position in the items array.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


### getItem(id: string)
Gets a binding object by ID.

#### Syntax
```js
bindingCollectionObject.getItem(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|ID of the binding object to be retrieved.|

#### Returns
[Binding](binding.md)

#### Examples

Create a table binding to monitor data changes in the table. When data is changed, the background color of the table will change to orange.

```js
(function () {
	// Create myTable
	Excel.run(function (ctx) {
		var table = ctx.workbook.tables.add("Sheet1!A1:C4", true);
		table.name = "myTable";
		return ctx.sync().then(function () {
			console.log("MyTable is Created!");

			//Create a new table binding for myTable
			Office.context.document.bindings.addFromNamedItemAsync("myTable", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
				if (asyncResult.status == "failed") {
					console.log("Action failed with error: " + asyncResult.error.message);
				}
				else {
					// If successful, add the event handler to the table binding.
					Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
				}
			});
		})
		.catch(function (error) {
			console.log(JSON.stringify(error));
		});
	});
	
	// When data in the table is changed, this event is triggered.
	function onBindingDataChanged(eventArgs) {
		Excel.run(function (ctx) {
			// Highlight the table in orange to indicate data changed.
			var fill = ctx.workbook.tables.getItem("myTable").getDataBodyRange().format.fill;
			fill.load("color");
			return ctx.sync().then(function () {
				if (fill.color != "Orange") {
					ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
 
					console.log("The value in this table got changed!");
				}
				else
					
			})
				.then(ctx.sync)
			.catch(function (error) {
				console.log(JSON.stringify(error));
			});
		});
	} 
})();
 


```



#### Examples
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
	return ctx.sync().then(function() {
			console.log(binding.type); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getItemAt(index: number)
Gets a binding object based on its position in the items array.

#### Syntax
```js
bindingCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Binding](binding.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
	return ctx.sync().then(function() {
			console.log(binding.type); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### load(param: object)
Fills the proxy object created in the JavaScript, layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, accepts a [loadOption](loadoption.md) object.|

#### Returns
void
### Property access examples

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < bindings.items.length; i++)
		{
			console.log(bindings.items[i].id);
			console.log(bindings.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Get the number of bindings.

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('count');
	return ctx.sync().then(function() {
		console.log("Bindings: Count= " + bindings.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
