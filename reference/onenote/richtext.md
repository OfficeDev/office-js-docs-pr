# RichText Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents a RichText object in a Paragraph.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the RichText object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-id)|
|text|string|Gets the text content of the RichText object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-text)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the RichText object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-paragraph)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-load)|

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

**id and text**
```js
OneNote.run(function (context) {

	// Get the collection of pageContent items from the page.
	var pageContents = context.application.getActivePage().contents;

	// Get the first PageContent on the page, and then get its outline's paragraphs.
	var outlinePageContents = [];
	var paragraphs = [];
	var richTextParagraphs = [];
	// Queue a command to load the id and type of each page content in the outline.
	pageContents.load("id,type");

	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {
			// Load all page contents of type Outline
			$.each(pageContents.items, function(index, pageContent) {
				if(pageContent.type == 'Outline')
				{
					pageContent.load('outline,outline/paragraphs,outline/paragraphs/type');
					outlinePageContents.push(pageContent);
				}
			});
			return context.sync();
		})
		.then(function () {
			// Load all rich text paragraphs across outlines
			$.each(outlinePageContents, function(index, outlinePageContent) {
				var outline = outlinePageContent.outline;
				paragraphs = paragraphs.concat(outline.paragraphs.items);
			});
			$.each(paragraphs, function(index, paragraph) {
				if(paragraph.type == 'RichText')
				{
					richTextParagraphs.push(paragraph);
					paragraph.load("id,richText/text");
				}
			});
			return context.sync();
		})
		.then(function () {
			// Display all rich text paragraphs to the console
			$.each(richTextParagraphs, function(index, richTextParagraph) {
				var richText = richTextParagraph.richText;
				console.log("Paragraph found with richtext content : " + richText.text + " and richtext id : " + richText.id);
			});
			return context.sync();
		});
});
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
}); 
```
