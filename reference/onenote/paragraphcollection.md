# ParagraphCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a collection of Paragraph objects.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of paragraphs in the page. Read-only.|
|items|[Paragraph[]](paragraph.md)|A collection of paragraph objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Paragraph](paragraph.md)|Gets a Paragraph object by ID or by its index in the collection. Read-only.|
|[getItemAt(index: number)](#getitematindex-number)|[Paragraph](paragraph.md)|Gets a paragraph on its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getItem(index: number or string)
Gets a Paragraph object by ID or by its index in the collection. Read-only.

#### Syntax
```js
paragraphCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|The ID of the Paragraph object, or the index location of the Paragraph object in the collection.|

#### Returns
[Paragraph](paragraph.md)

#### Examples
```js
OneNote.run(function (context) {

	// Get the collection of pageContent items from the page.
	var pageContents = context.application.getActivePage().contents;

	// Get the first PageContent on the page, and then get its Outline's first paragraph.
	var pageContent = pageContents.getItemAt(0);
	var paragraphs = pageContent.outline.paragraphs;

	var firstParagraph = paragraphs.getItemAt(0);

	// Queue a command to load the type and richText.text property of this paragraph.
	firstParagraph.load("id,type");


	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {
			// Write text from paragraph to console
			console.log("First Paragraph found with id : " + firstParagraph.id + " and type " + firstParagraph.type);
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
}); 
```
### getItemAt(index: number)
Gets a paragraph on its position in the collection.

#### Syntax
```js
paragraphCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Paragraph](paragraph.md)

#### Examples
```js
OneNote.run(function (context) {

	// Get the collection of pageContent items from the page.
	var pageContents = context.application.getActivePage().contents;

	// Get the first PageContent on the page, and then get its Outline's first paragraph.
	var pageContent = pageContents.getItemAt(0);
	var paragraphs = pageContent.outline.paragraphs;

	var firstParagraph = paragraphs.getItemAt(0);

	// Queue a command to load the type and richText.text property of this paragraph.
	firstParagraph.load("id,type");


	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {
			// Write text from paragraph to console
			console.log("First Paragraph found with id : " + firstParagraph.id + " and type " + firstParagraph.type);
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
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
### Property access examples

**items**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page, and then get its Outline's first paragraph.
    var pageContent = pageContents.getItem(0);
    var paragraphs = pageContent.outline.paragraphs;
	
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
			var firstParagraph = paragraphs.items[0];
            // Write text from first paragraph to console
			console.log("First Paragraph found with id : " + firstParagraph.id + " and type " + firstParagraph.type);
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

**traverse for richText**
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
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

