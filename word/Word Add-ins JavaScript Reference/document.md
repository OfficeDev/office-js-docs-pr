# Document 
The Document object is the top level object. A Document object contains one or more 
sections, content controls, and the body that contains the contents of the document.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|body|  [Body](body.md)   |Gets the body of the document. | 
|saved|  bool | Indicates whether the document has been changed. A value of true indicates that the document hasn't changed since it was last changed. | 



## Relationships

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|contentControls| [contentControlCollection](contentControlCollection.md)  | Gets the collection of content control objects that are in the current document. This includes content controls in the header, footer, and in the body of the document. | 
|[sections](#sections)| [SectionCollection](sectionCollection.md) | Gets the collection of section objects that are in the document. |


## Methods


| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getSelection()](#getselection())| [Range](range.md) |Gets the current selection of the document. Multiple selections are not supported. |
|[load(param: option)](#loadparam-option)| Document | Fills the document proxy object created in the JavaScript layer with property and object values specified in the parameter.  |
|[save()](#save)| void | Saves the document. This will use the Word default file naming convention if the document has not been saved before. |     

## API Specification

### getSelection()

Gets the current selection of the document. Multiple selections are not supported.

#### Syntax

```js
    document.getSelection();
```

#### Parameters

None

#### Returns

[Range](range.md)

[Back](#methods)

### load(param: option)

Fills the document proxy object created in the JavaScript layer with property and object values specified in the parameter.

#### Syntax

```js
    document.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns

void


[Back](#methods)



### save()

Saves the current document. 

#### Syntax

```js
    document.save();
```

#### Parameters 

None

#### Returns

void

#### Examples

```js
    var ctx = new Word.RequestContext();
    ctx.document.save();
```
[Back](#methods)




### Getter and Setter Examples

#### contentControls

```js
    // Gets a set of content control by tags and logs its content.
    var ctx = new Word.RequestContext();
    var ccs = ctx.document.contentControls;
    ctx.load(ccs,{select:'text'});

    ctx.executeAsync().then(
         function () {
             for(var i=0; i< ccs.items.length;i++)
            console.log( ccs.items[i].text);


         },
         function (result) {
             console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
             console.log(result.traceMessages);
         }
    );

```
[Back](#relationships)

#### sections

```js
    // Gets the paragraphs of the first section in the document.

    var ctx = new Word.RequestContext();

    var mySections = ctx.document.sections;
    ctx.load(mySections);

    var paras = mySections.getItem(0).body.paragraphs;
    ctx.load(paras);


    ctx.executeAsync().then(
        function () {
            var results = new Array();
            for (var i = 0; i < paras.items.length; i++) {
              console.log(paras.items[0].text);
            }  
        },
        function (result) {
            console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
            console.log(result.traceMessages);
        }
    );
```
[Back](#relationships)

