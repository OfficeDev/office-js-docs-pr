# ContentControl

Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. Currently, only rich text content controls are supported. 


## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|appearance|  string |Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'. |
|cannotDelete|  bool |Gets or sets a value that indicates whether the user can delete a content control from the active document.|
|cannotEdit|  bool | Gets or sets a value that indicates whether the user can edit the contents of a content control.|
|color|  string |   Gets or sets the color of the content control. Color is set in "#FFFFFF" format or by using the color name.|
|font|  [Font](font.md) | Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. |
|id|  string |Gets a string that represents the content control identifier. |
|parentContentControl|  [ContentControl](contentControl.md)   |Gets the content control that contains the content control. Returns null if there isn't a parent content control.|
|placeholderText|  string   | Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the  content control is empty.|
|removeWhenEdited|  bool | Gets or sets a value that indicates whether the content control is removed after it is edited.|
|title|  string  |  Gets or sets the title for a content control.   | 
|text|  string  |  Gets the text of the content control. |
|type|  string  | Gets or sets the content control type. Only rich text content controls are supported|
|style| string |Gets or sets the style used for the content control. This is the name of the pre-installed or custom style.|
|tag| string |Gets or sets a value to identify a content control. |



## Relationships

| Relationship     | Type    |Description|
|:-----------------|:--------|:----------|
|contentControls | [contentControlCollection](contentControlCollection.md)  | Gets the collection of content control objects in the current content control. | 
|inlinePictures | [inlinePictureCollection](inlinePictureCollection.md)  | Gets the collection of inlinePicture objects in the current content control. The collection does not include floating images.  | 
|paragraphs| [paragraphCollection](paragraphCollection.md)  | Get the collection of paragraph objects in the content control. |      

       

## Methods


| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[clear()](#clear)| void | Clears the contents of the content control. The user can perform the undo operation on the cleared content. |
|[delete(keepContent: bool)](#deletekeepcontent-bool)| void  | Deletes the content control and its content from the document. If keepContent is set to true, the content is not deleted. | 
|[getHtml()](#gethtml)| string  | Gets the HTML representation  of the content control object. | 
|[getOoxml()](#getooxml)| string  | Gets the Office Open XML (OOXML) representation  of the content control object. | 
|[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-insertlocation-string)| string |Inserts a document into the current content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.| 
|[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocation-string)| void | Inserts a break at the specified location. The insertLocation value can be 'Before', 'After', 'Start' or 'End'. | 
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)| [Paragraph](paragraph.md)  |Inserts a paragraph at the specified location. The insertLocation value can be 'Before', 'After', 'Start' or 'End'. | 
|[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-string)| [Range](range.md) | Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)| [Range](range.md)  |Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-string)| [Range](range.md)  |Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[load(param: option)](#loadparam-option)|void|Fills the content control proxy object created in the JavaScript layer with property and object values specified in the parameter.|
|[search(searchText : string, searchOptions: searchOptions)](#searchsearchtext-string-searchoptions-searchoptions)| [searchResultCollection](searchResultCollection.md) |Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects. | 
|[select()](#select())|  [Range](range.md) |Selects the content control. This causes Word to scroll to calling object.  | 
  
## API Specification


### clear()

Clears the content of the calling object.

#### Syntax
```js
    contentControl.clear();

```
#### Parameters

None

#### Returns

void

[Back](#methods)


### delete(keepContent: bool)
Deletes the content control and its content from the document. If keepContent is set to true, the content is not deleted.

#### Syntax
```js
    contentControl.Delete(keepContent: bool);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|keepContent|bool|Inidcates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted. |

### getHtml()

If keepContent is set to true, the content is not deleted.

#### Syntax
```js
    contentControl.getHtml();
```
#### Parameters

None

#### Returns

[Range](range.md).


[Back](#methods)

### getOoxml

Gets the Office Open XML (OOXML) representation  of the content control object. | 

#### Syntax
```js
    contentControl.getOoxml();
```
#### Parameters

None

#### Returns

[Range](range.md).

[Back](#methods)

### insertText(text: string, insertLocation: string)

Inserts text into the content control at the specified location.

#### Syntax
```js
    contentControl.insertText(text, insertLocation);
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`text`          | string | Required. The text to be inserted in to the content control.
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'. 

#### Returns

[Range](range.md).

[Back](#methods)

### insertHtml(html: string, insertLocation: string)

Inserts HTML into the content control at the specified location.

#### Syntax
```js
    contentControl.insertHtml(html, insertLocation);
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`html`          | string | Required. The HTML to be inserted in to the content control.
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'. 

#### Returns

[Range](range.md)



[Back](#methods)

### insertOoxml(ooxml: string, insertLocation: string)

Inserts OOXML into the content control at the specified location. 

#### Syntax
```js
    contentControl.insertOoxml(ooxml, insertLocation);
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`ooxml`          | string | Required. The OOXML to be inserted in to the content control. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'. |
 
#### Returns

[Range](range.md)


[Back](#methods)

### insertParagraph(paragraphText: string, insertLocation: string)

Inserts a paragraph at the specified location. The insertLocation value can be 'Before', 'After', 'Start' or 'End'. 

#### Syntax
```js
    contentControl.insertParagraph(paragraphText, insertLocation);
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`paragraphText`          | string | Paragrph text. null for blank Paragraph.|
`insertLocation`          | string | The value can be 'Before', 'After', 'Start' or 'End'. |


#### Returns

[Paragraph](Paragraph.md).


[Back](#methods)

### insertFileFromBase64(base64File: string, insertLocation: string)

Inserts a document into the current content control at the specified location.

#### Syntax
```js
    contentControl.insertFileFromBase64(base64File, insertLocation)
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`base64File`          | string | Required. Base64 encoded contents of the file to be inserted. 
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'. |


#### Returns

[Range](range.md)


[Back](#methods)

### insertBreak(breakType: string, insertLocation: string)

Inserts a break at the specified location.

#### Syntax
```js
    contentControl.insertBreak(breakType, insertLocation);
```
#### Parameters

Parameter      | Type   | Description
-------------- | ------ | ------------
`breakType`    | string | Required.  [Type of break](breakType.md)
`insertLocation` | string | The value can be 'Before', 'After', 'Start' or 'End'. |


#### Returns

[Range](range.md) collection.

[Back](#methods)


### load(param: option)

Fills the content control proxy object created in the JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
    contentControl.load(param);
```
#### Parameters

| Parameter      | Type   | Description
|  ------------- | ------ | ------------
|`param`         | object | A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns

void

[Back](#methods)

### search(searchText : string, searchOptions: searchOptions)

Performs a search with the specified search options on the scope of the content control object.

#### Syntax
```js
    contentControl.search(text, searchOptions);
```
#### Parameters

| Parameter      | Type   | Description
|  ------------- | ------ | ------------
|`text`          | string | Required. Text to be searched. |
|`searchOptions` | string |  |

#### Returns

[searchResultCollection](searchResultCollection.md) that contains range objects.


[Back](#methods)


### select()

Selects the content control content.  



#### Parameters

None

#### Returns

 [Range](range.md)

#### Examples

```js
    //Search and selects the first occurrence

    var ctx = new Word.RequestContext();
    var options = Word.SearchOptions.newObject(ctx);

    options.matchCase = false

    var results = ctx.document.body.search("Video", options);
    ctx.load(results, {select:"text, font/color", expand:"font"});
    ctx.references.add(results);

    ctx.executeAsync().then(
      function () {
        console.log("Found count: " + results.items.length + " " + results.items[0].font.color );
        for (var i = 0; i < results.items.length; i++) {
          results.items[i].font.color = "#FF0000"    // Change color to Red
          results.items[i].font.highlightColor = "#FFFF00";
          results.items[i].font.bold = true;
          if (i == 0)
            results.items[i].select();
        }
        ctx.references.remove(results);
        ctx.executeAsync().then(
          function () {
            console.log("Deleted");
          }
        );
      }
    );

```
[Back](#methods)