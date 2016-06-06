# Image Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|description|string|Gets or sets the description of the Image.|
|height|double|Gets or sets the height of the Image layout.|
|hyperlink|string|Gets or sets the hyperlink of the Image.|
|id|string|Gets the ID of the Image object. Read-only.|
|width|double|Gets or sets the width of the Image layout.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|pageContent|[PageContent](pagecontent.md)|Gets the PageContent object that contains the Image. Returns null if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page. Read-only.|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the Image. Returns null if the Image is not a direct child of a Paragraph. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getBase64Image()](#getbase64image)|string|Gets the base64-encoded binary representation of the Image.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getBase64Image()
Gets the base64-encoded binary representation of the Image.

#### Syntax
```js
imageObject.getBase64Image();
```

#### Parameters
None

#### Returns
string

#### Examples
```js

var image = null;
var imageString;

OneNote.run(function(ctx){
	// Get the current outline.			
	var outline = ctx.application.getActiveOutline();
	
	// Queue a command to load paragraphs and their types. 
	outline.load("paragraphs/type")
	return ctx.sync().
		then(function(){
			for (var i=0; i < outline.paragraphs.items.length; i++)
			{
				var paragraph = outline.paragraphs.items[i];
				if (paragraph.type == "Image")
				{
					image = paragraph.image;
				}
			}
		})
		.then(function(){
			if (image != null)
			{
				imageString = image.getBase64Image();
				return ctx.sync();
			}
		})
		.then(function(){
			console.log(imageString);
		});
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
