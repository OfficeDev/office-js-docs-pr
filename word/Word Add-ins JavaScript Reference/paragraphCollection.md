# ParagraphCollection

Contains a collection of [Paragraph](paragraph.md) objects. 

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|items|  array | Gets an array of paragraph objects. |


## Relationships
None  

## Methods

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)| [Paragraph](paragraph.md)   | Gets a paragraph object by its index in the collection. |
|[load(param: option)](#loadparam-option)|void|Fills the paragraph collection proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## API Specification

### getItem(index: number)

Gets a paragraph object by its index in the collection.

#### Syntax
```js
    paragraphCollection.getItem(index);
```
#### Parameters

| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|  A number that identifies the index location of a paragraph object.  |

#### Returns

[Paragraph](paragraph.md)  


[Back](#methods)


### load(param: option)
Fills the paragraph collection proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    paragraphCollection.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)



### Getter and setter examples

#### Get the text of each paragraph in the collection
```js
    // Iterate through all of the paragraphs in the documents and
    // report back the length and text of each paragraph.
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



