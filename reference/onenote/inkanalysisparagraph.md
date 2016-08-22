# InkAnalysisParagraph Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents ink analysis data for an identified paragraph formed by ink strokes.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the InkAnalysisParagraph object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-id)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|inkAnalysis|[InkAnalysis](inkanalysis.md)|Reference to the parent InkAnalysisPage. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-inkAnalysis)|
|lines|[InkAnalysisLineCollection](inkanalysislinecollection.md)|Gets the ink analysis lines in this ink analysis paragraph. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-lines)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-load)|

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

**lines**
```js
OneNote.run(function (ctx) {		
	var app = ctx.application;
	
	// Gets the active page.
	var page = app.getActivePage();
	
	// Load a line of ink words.
	page.load('inkAnalysisOrNull/paragraphs/lines');
	
	return ctx.sync()
		.then(function() {
			var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
			
			// Log id of each line in ink paragraphs.
			$.each(inkParagraphs.items, function(i, inkParagraph){
				var inkLines = inkParagraph.lines;
				$.each(inkLines.items, function (j, inkLine) {
					console.log(inkLine.id);
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