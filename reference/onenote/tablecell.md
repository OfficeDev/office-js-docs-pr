# TableCell Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a cell in a OneNote table.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|cellIndex|int|Gets the index of the cell in its row. Read-only.|
|id|string|Gets the ID of the cell. Read-only.|
|rowIndex|int|Gets the index of the cell's row in the table. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of Paragraph objects in the TableCell. Read-only.|
|parentRow|[TableRow](tablerow.md)|Gets the parent row of the cell. Read-only.|
|parentTable|[Table](table.md)|Gets the parent table of the cell. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Adds the specified HTML to the bottom of the TableCell.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### appendHtml(html: string)
Adds the specified HTML to the bottom of the TableCell.

#### Syntax
```js
tableCellObject.appendHtml(html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|The HTML string to append. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.|

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
