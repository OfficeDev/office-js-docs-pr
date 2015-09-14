# Paragraph
Represents a single paragraph in a selection, range, document, or document body.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|parentContentControl|  [ContentControl](contentControl.md) | Gets the content control that contains the paragraph. Returns null if there isn't a parent content control.|
|font|  [Font](font.md) | Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. |
|alignment| string |Gets or sets the alignment for a paragraph. The value can  be "left", "centered", "right", or "justified". |
|firstLineIndent| number |Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.|
|leftIndent| number | Gets or sets the left indent value, in points, for the paragraph.|
|lineSpacing| number | Gets or sets the line spacing, (in points) for the specified paragraph. In the Word UI, this value is divided by 12. |
|lineUnitAfter| number |Gets or sets the amount of spacing, in grid lines. after the paragraph.|
|lineUnitBefore| number |Gets or sets the amount of spacing, in grid lines, before the paragraph.
|outlineLevel| number |Gets or sets the outline level for the paragraph.
|rightIndent| number |Gets or sets the right indent value, in points, for the paragraph.
|spaceAfter| number |Gets or sets the spacing, in points, after the paragraph. |
|spaceBefore| number |Gets or sets the spacing, in points, before the paragraph. |
|style| string |Gets or sets the style used for the paragraph. This is the name of the pre-installed or custom style.|
|text|  string  |  Gets the text of the paragraph.  | 


## Relationships

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|contentControls| [contentControlCollection](contentControlCollection.md)  | Gets the collection of content control objects that are in the paragraph.|
|inlinepictures| [inlinePictureCollection](inlinePictureCollection.md)  |Gets the collection of inlinePicture objects that are in the paragraph. The collection does not include floating images.  | 


## Methods


| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[clear()](#clear)| void | Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content. | 
|[delete()](#delete)| void | Deletes the paragraph and its content from the document. |
|[getHtml()](#gethtml)| string  |  Gets the HTML representation  of the paragraph object. | 
|[getOoxml()](#getooxml)| string  | Gets the Office Open XML (OOXML) representation  of the paragraph object. | 
|[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocation-string)| void | Inserts a break at the specified location. The insertLocation value can be 'Start' or 'End'. | 
|[insertContentControl()](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the paragraph object with a Rich Text content control.| 
|[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-insertlocation-string)| [Range](range.md) |Inserts a document into the current paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.| 
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)| [Range](range.md)  | Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-string)| [InlinePicture](inlinePicture.md)  |Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Before', 'After', 'Start' or 'End'.| 
|[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-string)| [Range](range.md)  |Inserts OOXML into the paragraph at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)| [Paragraph](paragraph.md)  |Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'. | 
|[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-string)| [Range](range.md) | Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[load(param: option)](#loadparam-option)|void|Fills the paragraph proxy object created in the JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: searchOptions)](#searchsearchtext-string-searchoptions-searchoptions)| [searchResultCollection](searchResultCollection.md) |Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects. | 
|[select()](#select)|  [Range](range.md)   | Selects and navigates the Word UI to the paragraph. |

## API Specification

### clear()

Clears the contents of the paragraph object.

#### Syntax
```js
    paragraph.clear();
```
#### Parameters

None

#### Returns

void

[Back](#methods)

### delete()

Deletes the paragraph object.

#### Syntax
```js
    paragraph.delete();
```
#### Parameters

None

#### Returns

void

[Back](#methods)


### getHtml()

Gets the HTML representation  of the paragraph object.

#### Syntax
```js
    paragraph.getHtml();
```
#### Parameters

None

#### Returns

string

[Back](#methods)

### getOoxml()

Gets the Office Open XML (OOXML) representation  of the paragraph object.

#### Syntax
```js
    paragraph.getOoxml();
```
#### Parameters

None

#### Returns

string

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
`breakType`          | string | Required. The break type to add to the document. |
`insertLocation`          | string |  The value can be 'Before' or 'After'.|


#### Returns

void


[Back](#methods)


### insertContentControl()

Wraps the paragraph object with a Rich Text content control.

#### Syntax
```js
    paragraph.insertContentControl();
```
#### Parameters

None

#### Returns

[ContentControl](contentControl.md).

[Back](#methods)


### insertFileFromBase64(base64File: string, insertLocation: string)

Inserts a document into the current paragraph at the specified location.

#### Syntax
```js
    paragraph.insertFileFromBase64(base64File, insertLocation)
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`base64File`          | string | Required. The file base64 encoded file contents to be inserted. |
`insertLocation`          | string | The value can be 'Start' or 'End'.|



#### Returns

[Range](range.md)


[Back](#methods)




### insertHtml(html: string, insertLocation: string)

Inserts HTML into the paragraph at the specified location.

#### Syntax
```js
    paragraph.insertHtml(html, insertLocation);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`html`          | string | Required. The HTML to be inserted in the paragraph. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|

#### Returns

[Range](range.md)

[Back](#methods)


### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)

Inserts a picture into the paragraph at the specified location.

#### Syntax
```js
    paragraph.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`base64EncodedImage`          | string | Required. The HTML to be inserted in the paragraph. |
`insertLocation`          | string | The value can be 'Before', 'After', 'Start' or 'End'.|

#### Returns

[InlinePicture](inlinePicture.md)

[Back](#methods)



### insertOoxml(ooxml: string, insertLocation: string)

Inserts OOXML into the paragraph at the specified location. 

#### Syntax
```js
    paragraph.insertOoxml(ooxml, insertLocation);
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`ooxml`          | string | Required. The OOXML to be inserted in the paragraph. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|
 
#### Returns

[Range](range.md)

[Back](#methods)

### insertParagraph(paragraphText: string, insertLocation: string)

Inserts a paragraph at the specified location.

#### Syntax
```js
    paragraph.insertParagraph(paragraphText, insertLocation);
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

Inserts text into the paragraph at the specified location.

#### Syntax
```js
    paragraph.insertText(text, insertLocation);
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
Fills the paragraph proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    paragraph.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)




### search(searchText: string, searchOptions: searchOptions)

Executes a search on the scope of the paragraph object.

#### Syntax
```js
    paragraph.search(searchText, searchOptions)
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

Selects and navigates the Word UI to the paragraph.

#### Syntax
```js
    paragraph.select();
```
#### Parameters

None

#### Returns

 [Range](range.md) 


[Back](#methods)

### Getter and setter examples

#### Setting paragraph properties 

```js
    // Check out how it modifies your first paragraph's settings
    var ctx = new Word.RequestContext();
    var paras = ctx.document.body.paragraphs;
    ctx.load(paras, {select:"text"});
    ctx.references.add(paras);


    ctx.executeAsync().then(
      function () {
    var par = paras.items[0];
    par.lineSpacing = 45;
    par.alignment = "justified";
    par.spaceAfter = 45;
    par.firstLineIndent = 1;
    par.leftIndent = 2;
    par.lineUnitAfter = 2;
    par.lineUnitBefore = 5;
    par.outlineLevel = 10;

     ctx.executeAsync().then(
          function () {
            console.log("Success!!!" + par.lineSpacing);
         ctx.references.remove(par);
          }
        );

        //console.log("Success! Setting paragraph line spacing to " + par.lineSpacing);
      },
      function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
      }
    );

```


