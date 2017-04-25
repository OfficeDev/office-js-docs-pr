# CommentCollection Object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the CommentCollection for a given Shape.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Comment[]](comment.md)|A collection of comment objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|int|Gets the number of Comments.|
|[getItem(key: string)](#getitemkey-string)|[Comment](comment.md)|Gets the Comment using its name.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getCount()
Gets the number of Comments.

#### Syntax
```js
CommentCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: string)
Gets the Comment using its name.

#### Syntax
```js
CommentCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Key is the name of the Comment to be retrieved.|

#### Returns
[Comment](comment.md)

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
```js
 Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	var shapeName = "Position Belt.41";
	var shape = activePage.shapes.getItem(shapeName);
	var shapecomments= shape.comments;
        shapecomments.load();
        return ctx.sync().then(function () {
       	  for(var i=0; i<shapecomments.items.length;i++)
		{
       	    	 var comment= shapecomments.items[i];
	   	 console.log("comment Author: " + comment.author);
	   	 console.log("Comment Text: " + comment.text);
	 	   console.log("Date " + comment.date);
		}
	 });
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
