# Page Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a OneNote page.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|clientUrl|string|The client url of the page. Read only Read-only.|
|id|string|Gets the ID of the page. Read-only.|
|pageLevel|int|Gets or sets the indentation level of the page.|
|title|string|Gets or sets the title of the page.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|contents|[PageContentCollection](pagecontentcollection.md)|The collection of PageContent objects on the page. Read only Read-only.|
|parentSection|[Section](section.md)|Gets the section that contains the page. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addOutline(left: double, top: double, html: String)](#addoutlineleft-double-top-double-html-string)|[Outline](outline.md)|Adds an Outline to the page at the specified position.|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|Inserts a new page before or after the current page.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### addOutline(left: double, top: double, html: String)
Adds an Outline to the page at the specified position.

#### Syntax
```js
pageObject.addOutline(left, top, html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|left|double|The left position of the top, left corner of the Outline.|
|top|double|The top position of the top, left corner of the Outline.|
|html|String|An HTML string that describes the visual presentation of the Outline. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.|

#### Returns
[Outline](outline.md)

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var page = context.application.getActivePage();

    // Queue a command to add an outline with given html. 
    var outline = page.addOutline(200, 200,
"<p>Images and a table below:</p> \
 <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\"> \
 <img src=\"http://imagenes.es.sftcdn.net/es/scrn/6653000/6653659/microsoft-onenote-2013-01-535x535.png\"> \
 <table> \
   <tr> \
     <td>Jill</td> \
     <td>Smith</td> \
     <td>50</td> \
   </tr> \
   <tr> \
     <td>Eve</td> \
     <td>Jackson</td> \
     <td>94</td> \
   </tr> \
 </table>"     
        );

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
    })
});
```


### insertPageAsSibling(location: string, title: string)
Inserts a new page before or after the current page.

#### Syntax
```js
pageObject.insertPageAsSibling(location, title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|location|string|The location of the new page relative to the current page.  Possible values are: Before, After|
|title|string|The title of the new page.|

#### Returns
[Page](page.md)

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var newPage = activePage.insertPageAsSibling("After", "Next Page");

    // Queue a command to load the newPage to access its data.
    context.load(newPage);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("page is created with title: " + newPage.title);
        })
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        })
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
