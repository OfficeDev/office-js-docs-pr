# Range  
Represents a contiguous area in a document.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|font|  [Font](font.md) | Gets the text format of the range. Use this to get and set font name, size, color, and other properties. |
|parentContentControl|  [ContentControl](contentControl.md) | Gets the content control that contains the range. Returns null if there isn't a parent content control.|
|style| string |Gets or sets the style used for the range. This is the name of the pre-installed or custom style.|
|text|  string  |  Gets the text of the range.  | 

## Relationships

| Relationship     | Type    |Description|
|:-----------------|:--------|:----------|
|contentControls| [contentControlCollection](contentControlCollection.md)  | Gets the collection of content control objects that are in the range.|
|paragraphs| [paragraphCollection](paragraphCollection.md)  | Gets the collection of paragraph objects that are in the range. |     
    


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[clear()](#clear)| void | Clears the contents of the range object. The user can perform the undo operation on the cleared content. | 
|[delete()](#delete)| void  |Deletes the range and its content from the document. |
|[getHtml()](#gethtml)| string  | Gets the HTML representation  of the range object. | 
|[getOoxml()](#getooxml)| string  | Gets the OOXML representation  of the range object. | 
|[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocation-string)| void | Inserts a break at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertContentControl()](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the range object with a Rich Text content control.| 
|[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-insertlocation-string)| [Range](range.md) |Inserts a document into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.| 
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)| [Range](range.md)  | Inserts HTML into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-string)| [Range](range.md)  |Inserts OOXML into the range at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)| [Paragraph](paragraph.md)  |Inserts a paragraph into the range at the specified location. The insertLocation value can be 'Start' or 'End'. | 
|[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-string)| [Range](range.md) | Inserts text into the range at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[load(param: option)](#loadparam-option)|void|Fills the range proxy object created in the JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: searchOptions)](#searchsearchtext-string-searchoptions-searchoptions)| [searchResultCollection](searchResultCollection.md) |Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects. | 
|[select()](#select)| [Range](range.md)  | Selects and navigates the Word UI to the range. |


## API Specification

### clear()

Clears the content of the range object.

#### Syntax
```js
    range.clear();

```
#### Parameters

None

#### Returns

void

[Back](#methods)



### delete()

Deletes the range object.

#### Syntax
```js
    range.delete();
```
#### Parameters

None

#### Returns

void

[Back](#methods)

### getHtml()

Gets the HTML representation of the range object.

#### Syntax
```js
    range.getHtml();
```
#### Parameters

None

#### Returns

ClientResult


[Back](#methods)

### getOoxml()

Gets the Office Open XML (OOXML) representation of the range object.

#### Syntax
```js
    range.getOoxml();
```
#### Parameters

None

#### Returns

ClientResult

[Back](#methods)

### insertBreak(breakType: string, insertLocation: string)

Inserts a break at the specified location. 

#### Syntax
```js
    paragraph.insertBreak(breakType, insertLocation);
```
#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`breakType`          | string | Required. The break type to add to the range. |
`insertLocation`          | string |  The value can be 'Replace', 'Before' or 'After'.|


#### Returns

void


[Back](#methods)

### insertContentControl()

Wraps the range object with a Rich Text content control.

#### Syntax
```js
    range.insertContentControl();
```
#### Parameters

None

#### Returns

[ContentControl](contentControl.md).

#### Example

```js

    // Wraps the rnage object with a content control, then sets a few properties.
    var ctx = new Word.RequestContext();
    var range = ctx.document.getSelection();

    var myContentControl = range.insertContentControl();
    myContentControl.tag = "Customer-Address";
    myContentControl.title = "Enter Customer Address Here:";
    myContentControl.style = "Heading 1";
    myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
    myContentControl.cannotEdit = true;
    myContentControl.appearance = "tags";

    ctx.executeAsync().then(
      function () {
        console.log("Content control Id: " + myContentControl.id);
      },
      function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
      }
    );

```
[Back](#methods)


### insertFileFromBase64(base64File: string, insertLocation: string)

Inserts a file into the range at the specified location. 

#### Syntax
```js
    range.insertFileFromBase64(base64File, insertLocation)
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`base64File`          | string | Required. The file base64 encoded file contents to be inserted. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|



#### Returns

[Range](range.md)



[Back](#methods)


### insertHtml(html: string, insertLocation: string)

Inserts HTML into the range at the specified location. 

#### Syntax
```js
    range.insertHtml(html, insertLocation);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`html`          | string | Required. The HTML to be inserted in the range. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|

#### Returns

[Range](range.md)


[Back](#methods)

### insertOoxml(ooxml: string, insertLocation: string)

Inserts OOXML into the range at the specified location. 

#### Syntax
```js
    range.insertOoxml(ooxml, insertLocation);
```

#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`ooxml`          | string | Required. The OOXML to be inserted in the range. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|

#### Returns

[Range](range.md)

#### Example
```js
    var ctx = new Word.RequestContext();
    var range = ctx.document.getSelection();

    var ooxmlText =
      "<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'><w:r><w:rPr><w:b/><w:b-cs/><w:color w:val='FF0000'/><w:sz w:val='28'/><w:sz-cs w:val='28'/></w:rPr><w:t>Hello world (this should be bold, red, size 14).</w:t></w:r></w:p>";

    range.insertOoxml(ooxmlText, Word.InsertLocation.end);

    ctx.executeAsync().then(
       function () {
         console.log("Success");
       },
       function (result) {
         console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
         console.log(result.traceMessages);
       }
    );

```

[Back](#methods)

### insertParagraph(paragraphText: string, insertLocation: string)

Inserts a paragraph into the range at the specified location.

#### Syntax
```js
    range.insertParagraph(paragraphText, insertLocation);
```
#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`paragraphText`  | string | Required. The paragraph text to be inserted. Use null for inserting a blank paragraph.|
`insertLocation`  | string | The value can be 'Start' or 'End'.|


#### Returns

[Paragraph](Paragraph.md)

[Back](#methods)

### insertText(text: string, insertLocation: string)

Inserts text into the range at the specified location.

#### Syntax
```js
    range.insertText(text, insertLocation);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`text`          | string | Required. Text to be inserted. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|

#### Returns

[Range](range.md).

[Back](#methods)


### load(param: option)
Fills the range proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    range.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)


### search(searchText: string, searchOptions: searchOptions)

Executes a search on the scope of the range object.

#### Syntax
```js
    range.search(searchText, searchOptions)
```

#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`searchText`          | string | Required. The search text.|
`searchOptions` | [searchOptions](searchOptions.md) | Required. Options for the search.|

#### Returns

[SearchResultCollection](searchResultCollection.md)


[Back](#methods)


### select()

Selects and navigates the Word UI to the range.

#### Syntax
```js
    paragraph.select();
```
#### Parameters

None

#### Returns

[Range](range.md)


[Back](#methods)
