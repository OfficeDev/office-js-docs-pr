# Application Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|activeNotebook|[Notebook](notebook.md)|Gets the active notebook. Read-only.|
|activePage|[Page](page.md)|Gets the active page. Read-only.|
|activeSection|[Section](section.md)|Gets the active section. Read-only.|
|notebooks|[NotebookCollection](notebookcollection.md)|Gets the collection of notebooks that are open in the OneNote application instance. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getNotebookById(id: string)](#getnotebookbyidid-string)|[Notebook](notebook.md)|Gets a notebook by ID.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[navigateToPage(page: Page)](#navigatetopagepage-page)|void|Opens the specified page in the application instance.|

## Method Details


### getNotebookById(id: string)
Gets a notebook by ID.

#### Syntax
```js
applicationObject.getNotebookById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|The ID of the notebook.|

#### Returns
[Notebook](notebook.md)

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

### navigateToPage(page: Page)
Opens the specified page in the application instance.

#### Syntax
```js
applicationObject.navigateToPage(page);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|page|Page|The page to open.|

#### Returns
void
