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
|[insertSectionAsSibling(location: string, title: string)](#insertsectionassiblinglocation-string-title-string)|[Section](section.md)|Insert a new Section before or after the current section|
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

### insertSectionAsSibling(location: string, title: string)
Insert a new Section before or after the current section

#### Syntax
```js
sectionObject.insertSectionAsSibling(location, title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|location|string|Location of the new Section  Possible values are: Before, After|
|title|string|Title for the new Section|

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
