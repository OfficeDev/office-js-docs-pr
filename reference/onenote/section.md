# Section Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a OneNote section. Sections can contain pages.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the section. Read-only.|
|name|string|Gets the name of the section. Read-only.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|notebook|[Notebook](notebook.md)|Gets the notebook that contains the section. Read-only.|
|sectionGroup|[SectionGroup](sectiongroup.md)|Gets the section group that contains the section. Returns null if the section is a direct child of the notebook. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addPage(title: string)](#addpagetitle-string)|[Page](page.md)|Adds a new page to the end of the section.|
|[getPages()](#getpages)|[PageCollection](pagecollection.md)|Gets the collection of pages in the section.|
|[insertSectionAsSibling(location: string, title: string)](#insertsectionassiblinglocation-string-title-string)|[Section](section.md)|Inserts a new section before or after the current section.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### addPage(title: string)
Adds a new page to the end of the section.

#### Syntax
```js
sectionObject.addPage(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|string|The title of the new page.|

#### Returns
[Page](page.md)

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active section.
    var section = context.application.activeSection;

    // Queue a command to add a new page.
    var page = section.addPage("sample page");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("new page title is " + page.title);
        })
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```

### getPages()
Gets the collection of pages in the section.

#### Syntax
```js
sectionObject.getPages();
```

#### Parameters
None

#### Returns
[PageCollection](pagecollection.md)

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active section.
    var section = context.application.activeSection;

    // Queue a command to get pages of the section. 
    var pages = section.getPages();

    // Queue a command to load the pages. 
    context.load(pages);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(pages.items, function(index, page) {
                console.log("Page title: " + page.title);
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

### insertSectionAsSibling(location: string, title: string)
Inserts a new section before or after the current section.

#### Syntax
```js
sectionObject.insertSectionAsSibling(location, title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|location|string|The location of the new section relative to the current section.  Possible values are: Before, After|
|title|string|The name of the new section.|

#### Returns
[Section](section.md)

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
