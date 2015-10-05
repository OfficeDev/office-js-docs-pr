# ParagraphCollection object (JavaScript API for Word)

Contains a collection of [paragraph](paragraph.md) objects.

_Applies to: Word 2016_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Paragraph[]](paragraph.md)|A collection of paragraph objects. Read-only.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)|[Paragraph](paragraph.md)|Gets a paragraph object by its index in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method details

### getItem(index: number)
Gets a paragraph object by its index in the collection.

#### Syntax
```js
paragraphCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a paragraph object.|

#### Returns
[Paragraph](paragraph.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
	
	// Create a proxy object for the paragraphs collection.
	var paragraphs = context.document.body.paragraphs;
	
	// Queue a commmand to load the style property for all of the paragraphs.
	context.load(paragraphs, 'style');
	
	// Synchronize the document state by executing the queued commands, 
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		
		// Queue a command to get the first paragraph and create a 
		// proxy paragraph object.
		var paragraph = paragraphs.getItem(0); 
		
		// Queue a command to select the paragraph. The Word UI will 
		// move to the selected paragraph.
		paragraph.select();
		
		// Synchronize the document state by executing the queued commands, 
		// and return a promise to indicate task completion.
		return context.sync().then(function () {
			console.log('Selected the last paragraph.');
		});      
	});  
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the text and style properties for all of the paragraphs.
    context.load(paragraphs, 'text, style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the last paragraph and create a 
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1]; 
        
        // Queue a command to select the paragraph. The Word UI will 
        // move to the selected paragraph.
        paragraph.select();
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```