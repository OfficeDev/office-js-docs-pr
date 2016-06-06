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
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the type and richText.text property of each paragraph.
    paragraphs.load("type,richText/text");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                if (paragraph.type === 'RichText') { 
                    console.log("RichText text: " + paragraph.richText.text);
                }
            });
        })                
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }); 
    });
```