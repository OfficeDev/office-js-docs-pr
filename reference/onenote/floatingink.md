# FloatingInk Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents a group of ink strokes.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the FloatingInk object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-id)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|inkStrokes|[InkStrokeCollection](inkstrokecollection.md)|Gets the strokes of the FloatingInk object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-inkStrokes)|
|pageContent|[PageContent](pagecontent.md)|Gets the PageContent parent of the FloatingInk object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-pageContent)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-load)|

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

**id**
```js
OneNote.run(function(context) {

	// Gets the active page.
	var page = context.application.getActivePage();
	var contents = page.contents;
	
	// Load page contents and their types.
	page.load('contents/type');
	return context.sync()
		.then(function(){
		
			// Load every ink content.
			$.each(contents.items, function(i, content) {
				if (content.type == "Ink")
				{
					content.load('ink/id');
				}							
			})
			return context.sync();
		})
		.then(function(){
		
			// Log ID of every ink content.
			$.each(contents.items, function(i, content) {
				if (content.type == "Ink")
				{
					console.log(content.ink.id);
				}							
			})				
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
}); 
```
