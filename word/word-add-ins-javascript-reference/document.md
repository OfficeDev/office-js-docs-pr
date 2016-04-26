# Document Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|saved|bool|Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.|WordApi1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|body|[Body](body.md)|Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.|WordApi1.1||
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the current document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.|WordApi1.1||
|sections|[SectionCollection](sectioncollection.md)|Gets the collection of section objects in the document. Read-only.|WordApi1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getSelection()](#getselection)|[Range](range.md)|Gets the current selection of the document. Multiple selections are not supported.|WordApi1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|WordApi1.1|
|[open()](#open)|void|Open the document.|WordApiWordApiDesktop, 1.3|
|[save()](#save)|void|Saves the document. This will use the Word default file naming convention if the document has not been saved before.|WordApi1.1|

## Method Details


### getSelection()
Gets the current selection of the document. Multiple selections are not supported.

#### Syntax
```js
documentObject.getSelection();
```

#### Parameters
None

#### Returns
[Range](range.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### open()
Open the document.

#### Syntax
```js
documentObject.open();
```

#### Parameters
None

#### Returns
void

### save()
Saves the document. This will use the Word default file naming convention if the document has not been saved before.

#### Syntax
```js
documentObject.save();
```

#### Parameters
None

#### Returns
void
