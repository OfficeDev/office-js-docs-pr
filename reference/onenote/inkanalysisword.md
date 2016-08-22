# InkAnalysisWord Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents ink analysis data for an identified word formed by ink strokes.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the InkAnalysisWord object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-id)|
|languageId|string|The id of the recognized language in this inkAnalysisWord. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-languageId)|
|wordAlternates|string|The words that were recognized in this ink word, in order of likelihood. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-wordAlternates)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|line|[InkAnalysisLine](inkanalysisline.md)|Reference to the parent InkAnalysisLine. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-line)|
|strokePointers|[InkStrokePointer](inkstrokepointer.md)|Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-strokePointers)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-load)|

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

**wordAlternates and languageId**
```js
OneNote.run(function (ctx) {		
	var app = ctx.application;
	
	// Gets the active page.
	var page = app.getActivePage();
	
	page.load('inkAnalysisOrNull/paragraphs/lines/words');
	return ctx.sync()
		.then(function() {
			var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
			$.each(inkParagraphs.items, function(i, inkParagraph) {
				var inkLines = inkParagraph.lines;
				$.each(inkLines.items, function(j, inkLine) {
					var inkWords = inkLine.words;
					$.each(inkWords.items, function(k, inkWord) {
					
						// Log language Id of the word
						console.log(inkWord.languageId);
						
						// Log every ink analyzed words.
						$.each(inkWord.wordAlternates, function(l, word) {
							console.log(word);									
						})
					})
				})
			})
		})
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
}); 
```