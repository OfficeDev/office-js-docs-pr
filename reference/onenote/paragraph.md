# Paragraph Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the Paragraph object. Read-only.|
|type|string|Gets the type of the Paragraph object. Read-only. Possible values are: RichText, Image, Table, Other.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|image|[Image](image.md)|Gets the Image object in the Paragraph. Returns null if ParagraphType is not Image. Read-only.|
|outline|[Outline](outline.md)|Gets the Outline object that contains the Paragraph. Read-only.|
|richText|[RichText](richtext.md)|Gets the RichText object in the Paragraph. Returns null if ParagraphType is not RichText. Read-only Read-only.|
|table|[Table](table.md)|Gets the Table object in the Paragraph. Returns null if ParagraphType is not Table. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the paragraph|
|[insertHtmlAsSibling(insertLocation: string, html: string)](#inserthtmlassiblinginsertlocation-string-html-string)|void|Inserts the specified HTML content|
|[insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)](#insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double)|[Image](image.md)|Inserts the image at the specified insert location..|
|[insertRichTextAsSibling(insertLocation: string, paragraphText: string)](#insertrichtextassiblinginsertlocation-string-paragraphtext-string)|[RichText](richtext.md)|Inserts the paragraph text at the specifiec insert location.|
|[insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])](#inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string)|[Table](table.md)|Adds a table with the specified number of rows and columns before or after the current paragraph.|
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

### insertHtmlAsSibling(insertLocation: string, html: string)
Inserts the specified HTML content

#### Syntax
```js
paragraphObject.insertHtmlAsSibling(insertLocation, html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|The location of new contents relative to the current Paragraph.  Possible values are: Before, After|
|html|string|An HTML string that describes the visual presentation of the content. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.|

#### Returns
void

### insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)
Inserts the image at the specified insert location..

#### Syntax
```js
paragraphObject.insertImageAsSibling(insertLocation, base64EncodedImage, width, height);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|The location of the table relative to the current Paragraph.  Possible values are: Before, After|
|base64EncodedImage|string|HTML string to append.|
|width|double|Optional. Width in the unit of Points. The default value is null and image width will be respected.|
|height|double|Optional. Height in the unit of Points. The default value is null and image height will be respected.|

#### Returns
[Image](image.md)

### insertRichTextAsSibling(insertLocation: string, paragraphText: string)
Inserts the paragraph text at the specifiec insert location.

#### Syntax
```js
paragraphObject.insertRichTextAsSibling(insertLocation, paragraphText);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|The location of the table relative to the current Paragraph.  Possible values are: Before, After|
|paragraphText|string|HTML string to append.|

#### Returns
[RichText](richtext.md)

### insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])
Adds a table with the specified number of rows and columns before or after the current paragraph.

#### Syntax
```js
paragraphObject.insertTableAsSibling(insertLocation, rowCount, columnCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|The location of the table relative to the current Paragraph.  Possible values are: Before, After|
|rowCount|number|The number of rows in the table.|
|columnCount|number|The number of columns in the table.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[Table](table.md)

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
    var pageContents = context.application.getActivePage().contents;
    
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