# Outline Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a container for Paragraph objects.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the Outline object. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|pageContent|[PageContent](pagecontent.md)|Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of Paragraph objects in the Outline. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Adds the specified HTML to the bottom of the Outline.|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|Adds the specified image to the bottom of the Outline.|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|Adds the specified text to the bottom of the Outline.|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|Adds a table with the specified number of rows and columns to the bottom of the outline.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[select()](#select)|void|Selects the Outline and bring it to the view.|

## Method Details


### appendHtml(html: string)
Adds the specified HTML to the bottom of the Outline.

#### Syntax
```js
outlineObject.appendHtml(html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|The HTML string to append. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.|

#### Returns
void

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendHtml("<p>new paragraph</p>");

                // Run the queued commands.
                return context.sync();
            }
        })
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```

### appendImage(base64EncodedImage: string, width: double, height: double)
Adds the specified image to the bottom of the Outline.

#### Syntax
```js
outlineObject.appendImage(base64EncodedImage, width, height);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|HTML string to append.|
|width|double|Optional. Width in the unit of Points. The default value is null and image width will be respected.|
|height|double|Optional. Height in the unit of Points. The default value is null and image height will be respected.|

#### Returns
[Image](image.md)

### appendRichText(paragraphText: string)
Adds the specified text to the bottom of the Outline.

#### Syntax
```js
outlineObject.appendRichText(paragraphText);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|HTML string to append.|

#### Returns
[RichText](richtext.md)

### appendTable(rowCount: number, columnCount: number, values: string[][])
Adds a table with the specified number of rows and columns to the bottom of the outline.

#### Syntax
```js
outlineObject.appendTable(rowCount, columnCount, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowCount|number|Required. The number of rows in the table.|
|columnCount|number|Required. The number of columns in the table.|
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
Selects the Outline and bring it to the view.

#### Syntax
```js
outlineObject.select();
```

#### Parameters
None

#### Returns
void
