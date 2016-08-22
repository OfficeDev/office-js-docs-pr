# ImageOcrData Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents data obtained by OCR (optical character recognition) of an image

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|ocrLanguageId|string|Represents the OCR language, with values such as EN-US|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-ocrLanguageId)|
|ocrText|string|Represents the text obtained by OCR of the image|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-ocrText)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-load)|

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
**ocrText and ocrLanguageId**
```js
var image = null;

OneNote.run(function(ctx){
	// Get the current outline.
	var outline = ctx.application.getActiveOutline();

	// Queue a command to load paragraphs and their types.
	outline.load("paragraphs")
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
			if (image != null)
			{
			   image.load("ocrData");
			}
			return ctx.sync();
		})
		.then(function(){
			
			// Log ocrText and ocrLanguageId
			console.log(image.ocrData.ocrText);
			console.log(image.ocrData.ocrLanguageId);
		});
}).catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```
