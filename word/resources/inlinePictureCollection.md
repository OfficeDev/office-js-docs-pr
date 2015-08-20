# InlinePictureCollection

Contains a collection of [InlinePicture](inlinePicture.md) objects. 


## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|items|  array | Gets an array of inline picture objects. |


## Relationships
None  

## Methods

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)| [InlinePicture](inlinePicture.md)   | Gets an inline picture object by its index in the collection. |
|[load(param: option)](#loadparam-option)|void|Fills the inline picture collection proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## API Specification

### getItem(index: number)

Gets an inline picture object by its index in the collection.

#### Syntax
```js
    inlinePictureCollection.getItem(index);
```
#### Parameters

| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|  A number that identifies the index location of an inline picture object.  |

#### Returns

[InlinePicture](inlinePicture.md)

#### Example

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
[Back](#methods)