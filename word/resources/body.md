# Body 
Represents the body of a document or a section.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|font|  [Font](font.md) | Gets the text format of the body. Use this to get and set font name, size, color, and other properties. |
|parentContentControl|  [ContentControl](contentControl.md) | Gets the content control that contains the body. Returns null if there isn't a parent content control.|
|style| string | Gets or sets the style used for the content control. This is the name of the pre-installed or custom style. |
|text| string | Gets the text of the content control.  Use the insertText method to insert text.  |

## Relationships

| Relationship     | Type    |Description|
|:-----------------|:--------|:----------|
|contentControls| [contentControlCollection](contentControlCollection.md)  | Gets the collection of content control objects that are in the current document body.|
|inlinepictures| [inlinePictureCollection](inlinePictureCollection.md)  |Gets the collection of inlinePicture objects that are in the document body. The collection does not include floating images.  | 
|paragraphs| [paragraphCollection](paragraphCollection.md)  | Gets the collection of paragraph objects that are in the document body. |      
    


## Methods

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[clear()](#clear)| void | Clears the contents of the body object. The user can perform the undo operation on the cleared content. | 
|[getHtml()](#gethtml)| string  | Gets the HTML representation  of the body object.| 
|[getOoxml()](#getooxml)| string  | Gets the Office Open XML (OOXML) representation of the body object. | 
|[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocation-string)| void | Inserts a break at the specified location. The insertLocation value can be 'Start' or 'End'. | 
|[insertContentControl()](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the body object with a Rich Text content control. |
|[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-insertlocation-string)| string |Inserts a document into the current document at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.| 
|[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-string)| [Range](range.md) | Inserts text into the document body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)| [Range](range.md)  |Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-string)| [Range](range.md)  |Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'. | 
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)| [Paragraph](paragraph.md)  |Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'. | 
|[load(param: option)](#loadparam-option)|void|Fills the body proxy object created in the JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: searchOptions)](#searchsearchtext-string-searchoptions-searchoptions)| [searchResultCollection](searchResultCollection.md) |Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects. | 
|[select()](#select)| [Range](range.md) |Selects the entire body of the document |

## API Specification


### clear()

Clears the contents of the body object.

#### Syntax
```js
    ctx.document.body.clear();
```
#### Parameters

None

#### Returns

void


#### Examples

```js

    //Clear content of the document body

    var ctx = new Word.RequestContext();

    ctx.document.body.clear();
    ctx.executeAsync().then(
       function () {
         console.log("Success!!");
       },
       function (result) {
         console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
         console.log(result.traceMessages);
       }
    );

```
[Back](#methods)

### getText

Gets the plain text value of the body object.

#### Syntax
```js
    myBody.text
```
#### Parameters

None

#### Returns

[Range](range.md).


#### Examples

```js

    //gets the text of the entire body.
    var ctx = new Word.RequestContext();
    var myBody = ctx.document.body
    ctx.load(myBody, {select:'text'});
    ctx.executeAsync().then(
        function () {
        console.log(myBody.text);    
        },
        function (result) {
            console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
            console.log(result.traceMessages);
        }
    );
```
[Back](#methods)

### getHtml

Gets the HTML representation of the body object.

#### Syntax
```js
    var myTHTML  = document.body.getHtml();
```
#### Parameters

None

#### Returns

[Range](range.md).


#### Examples

```js
    var myHTML  = document.body.getHtml();
```
[Back](#methods)

### getOoxml

Gets the Office Open XML (OOXML) representation of the body object.

#### Syntax
```js
    document.body.getOoxml();
```
#### Parameters

None

#### Returns

[Range](range.md).


#### Examples

```js
    var myOOXML  = document.body.getOoxml();
```
[Back](#methods)



### insertText(text: string, insertLocation: string)

Inserts text into the document body at the specified location. 

#### Syntax
```js
    document.body.insertText(text, insertLocation);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`text`          | string | Required. Text to be inserted. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|

#### Returns

[Range](range.md).


#### Examples

```js
    //get inserts some text at the end of the document.
    var ctx = new Word.RequestContext();
    ctx.document.body.insertText("new text", "end");
    ctx.executeAsync().then(
        function () {
        console.log("Success!!");    
        },
        function (result) {
            console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
            console.log(result.traceMessages);
        }
    );
```
[Back](#methods)

### insertHtml(html: string, insertLocation: string)

Inserts HTML into the document body at the specified location. 

#### Syntax
```js
    document.body.insertHtml(html, insertLocation);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`html`          | string | Required. The HTML to be inserted in the document. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|

#### Returns

[Range](range.md) .


#### Examples

```js
    //inserts some html at the end of the doc :) 
    var ctx = new Word.RequestContext();
    ctx.document.body.insertHtml("<b>This is some bold text</b>", "End");
    ctx.executeAsync().then(
        function () {
        console.log("Success!!");    
        },
        function (result) {
            console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
            console.log(result.traceMessages);
        }
    );
```
[Back](#methods)

### insertOoxml(ooxml: string, insertLocation: string)

Inserts OOXML into the document body at the specified location. 

#### Syntax

```js
    document.body.insertOoxml(ooxml, insertLocation);
```

#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`ooxml`          | string | Required. The OOXML to be inserted. |
`insertLocation` | string | The value can be 'Replace', 'Start' or 'End'.|

#### Returns

[Range](range.md) collection.

[Back](#methods)




### insertParagraph(paragraphText: string, insertLocation: string)

Inserts a paragraph into the document body at the specified location. 

#### Syntax
```js
    document.body.insertParagraph(paragraphText, insertLocation);
```
#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`paragraphText`  | string | Required. The paragraph text to be inserted. Use null for inserting a blank paragraph.|
`insertLocation`  | string | The value can be 'Start' or 'End'.|


#### Returns

[Paragraph](Paragraph.md).


#### Examples

```js

    //Inserting paragraphs at the end of the document.

    var ctx = new Word.RequestContext();

    var myPar = ctx.document.body.insertParagraph("Bibliography","end");
    myPar.style = "Heading 1";

    var myPar2 = ctx.document.body.insertParagraph("this is my first book","end");
    myPar2.style = "Normal"



    ctx.executeAsync().then(
         function () {
             console.log("Success!!");
         },
         function (result) {
             console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
            // console.log(result.traceMessages);
         }
    );


```
[Back](#methods)

### insertContentControl()

Wraps the body object with a Rich Text content control.

#### Syntax
```js
    document.body.insertContentControl();
```
#### Parameters

None

#### Returns

[ContentControl](contentControl.md).


#### Examples

```js

    // wraps the current selection with a content control, then sets a few properties.
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

### load(param: option)
Fills the body proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    document.body.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)

### search(searchText: string, searchOptions: searchOptions)

Performs a search with the specified searchOptions on the scope of the body object.

#### Syntax
```js
    document.body.search(searchText, searchOptions)
```

#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`searchText`          | string | Required. The seearch text.|
`searchOptions` | [searchOptions](searchOptions.md) | Required. Options for the search.|

#### Returns

[SearchResultCollection](searchResultCollection.md) collection.


#### Examples

```js

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
          if (i == 3)
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


### insertFileFromBase64(base64File:string, insertLocation:string)

Inserts a file into the document body at the specified location. 

#### Syntax
```js
    document.body.insertFileFromBase64(base64File, insertLocation)

```
#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`base64File`          | string | Required. The file base64 encoded file contents to be inserted. |
`insertLocation`          | string | The value can be 'Replace', 'Start' or 'End'.|


#### Returns

[Range](range.md)


[Back](#methods)



### insertBreak(breakType: string, insertLocation: string)

Inserts a break at the specified location.

#### Syntax
```js
    document.body.insertBreak(breakType, insertLocation);
```
#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`breakType`          | string | Required. The break type to add to the document. |
`insertLocation`          | string |  The value can be 'Start' or 'End'.|


#### Returns

[Range](range.md)


#### Examples

```js
    //inserts a page break and then adds a paragraph

    var ctx = new Word.RequestContext();

    ctx.document.body.insertBreak("page", "End");
    ctx.document.body.insertParagraph("Hello after break!","End");

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





### select()

Selects the body.  

#### Syntax
```js
    document.body.select();
```
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

### Getter and setter examples

#### contentControls

```javascript

    // enumerates all the content controls in the document
    var ctx = new Word.RequestContext();
    var cCtrls = ctx.document.body.contentControls;
    ctx.load(cCtrls,{select:'appearance,text'});  // just need these properties!

    ctx.executeAsync().then(
    function () {
            var results = new Array();

            for (var i = 0; i < cCtrls.items.length; i++) {
               console.log("contentControl[" + i + "].text = " + cCtrls.items[i].text + " Appearance:" +
                            cCtrls.items[i].appearance );
          }
            ctx.executeAsync().then(
                function () {
                   console.log("Success!!");
                }
            );
        },
        function (result) {
            console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
            console.log(result.traceMessages);
        }
    );


```
[Back](#relationships)


#### paragraphs 

```js

    // this example iterates all the paragraphs in the documents and reports back the 
    // length and text of each paragraph in the document
    var ctx = new Word.RequestContext();
    var paras = ctx.document.body.paragraphs;
    ctx.load(paras,{select:"text"});

    ctx.executeAsync().then(
      function () {
        for (var i = 0; i < paras.items.length; i++) {
          console.log("paras[" + i + "].content  = " + paras.items[i].text);
        }
      },
      function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
      }
    );
```
[Back](#relationships)


### inlinePictures 

```js

    //gets all the images in the body of the document and then gets the base64 for each.
    var ctx = new Word.RequestContext();

    var pics = ctx.document.body.inlinePictures;
    ctx.load(pics);
    ctx.references.add(pics);

    ctx.executeAsync().then(
      function () {
        var results = new Array();

        for (var i = 0; i < pics.items.length; i++) {
          results.push(pics.items[i].getBase64ImageSrc());
        }
        ctx.executeAsync().then(
          function () {
            for (var i = 0; i < results.length; i++) {
              console.log("pics[" + i + "].base64 = " + results[i].value);
            }
          }
        );
      },
      function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
      }
    );


```
[Back](#relationships)
