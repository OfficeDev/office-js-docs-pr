# Comment Object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the Comment.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|author|string|A string that specifies the name of the author of the comment.|
|text|string|A string that contains the comment text.|
|date|string|A string that specifies the date when the comment was created.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


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
