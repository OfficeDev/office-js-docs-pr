# ShapeDataItemCollection object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the ShapeDataItemCollection for a given Shape.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|items|[ShapeDataItem[]](shapedataitem.md)|A collection of shapeDataItem objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|int|Gets the number of Shape Data Items.|
|[getItem(key: string)](#getitemkey-string)|[ShapeDataItem](shapedataitem.md)|Gets the ShapeDataItem using its name.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getCount()
Gets the number of Shape Data Items.

#### Syntax
```js
shapeDataItemCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: string)
Gets the ShapeDataItem using its name.

#### Syntax
```js
shapeDataItemCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Key is the name of the ShapeDataItem to be retrieved.|

#### Returns
[ShapeDataItem](shapedataitem.md)

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
### Property access examples
```js
Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	var shape = activePage.shapes.getItem(0);
        var shapeDataItems = shape.shapeDataItems;
        shapeDataItems.load();
        return ctx.sync().then(function() {
            for (var i = 0; i < shapeDataItems.items.length; i++)
            {
                console.log(shapeDataItems.items[i].label);
                console.log(shapeDataItems.items[i].value);
            }
        });
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
