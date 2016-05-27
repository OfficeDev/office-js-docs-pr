# Paragraph Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the Paragraph object. Read-only.|
|type|string|Gets the type of the Paragraph object. Read-only. Possible values are: RichText, Image, Other.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|image|[Image](image.md)|Gets the Image object in the Paragraph. Returns null if ParagraphType is not Image. Read-only.|
|outline|[Outline](outline.md)|Gets the Outline object that contains the Paragraph. Read-only.|
|richText|[RichText](richtext.md)|Gets the RichText object in the Paragraph. Returns null if ParagraphType is not RichText. Read-only Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the paragraph|
|[insertAsSibling(html: string, insertLocation: string)](#insertassiblinghtml-string-insertlocation-string)|void|Inserts the specified HTML content|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[select()](#select)|void|Selects the paragraph|

## Method Details


### delete()
Deletes the paragraph

#### Syntax
```js
paragraphObject.delete();
```

#### Parameters
None

#### Returns
void

### insertAsSibling(html: string, insertLocation: string)
Inserts the specified HTML content

#### Syntax
```js
paragraphObject.insertAsSibling(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|An HTML string that describes the visual presentation of the content. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.|
|insertLocation|string|The location of new contents relative to the current Paragraph.  Possible values are: Before, After|

#### Returns
void

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

### select()
Selects the paragraph

#### Syntax
```js
paragraphObject.select();
```

#### Parameters
None

#### Returns
void
### Property access examples

**id and type**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().getContents();
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                console.log("Paragraph type: " + paragraph.type);
                console.log("Paragraph ID: " + paragraph.id);
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