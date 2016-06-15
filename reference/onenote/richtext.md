# RichText Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a RichText object in a Paragraph.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the RichText object. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-id)|
|text|string|Gets the text content of the RichText object. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-text)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the RichText object. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-paragraph)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-load)|

## Method Details


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
