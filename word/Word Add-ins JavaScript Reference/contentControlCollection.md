# ContentControlCollection

A collection of ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain content such as dates, lists, or paragraphs of formatted text.


## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`items`|  array | An array that contains content controls. 

## Relationships
None  

## Methods

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[getById(id: string)](#getbyidid-string)| [contentControl](contentControl.md) | Gets a content control by its identifier. | 
|[getByTag(tag: string )](#getbytagtag-string)| [contentControlCollection](contentControlCollection.md)  |Gets the content controls that have the specified tag. | 
|[getByTitle(title: string)](#getbytitletitle-string)| [contentControlCollection](contentControlCollection.md) |Gets the content controls that have the specified title. |   
|[getItem(index: number)](#getitemindex-number)| [contentControl](contentControl.md)   | Gets a content control by its index in the collection. |
|[load(param: option)](#loadparam-option)|void|Fills the content control collection proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## API Specification


### getById(id: string)

Gets a content control by its identifier.

#### Syntax
```js
    contentControlCollection.getById(id);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`id`          | string | A content control identifier. |

#### Returns

[contentControl](contentControl.md)


### getByTag(tag: string )
Gets the content controls that have the specified tag.


#### Syntax
```js
    contentControlCollection.getById(tag);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`tag`          | string | A tag set on a content control.|

#### Returns

[contentControlCollection](contentControlCollection.md)

#### Example
```js
    // gets Content control by tags and prints its content.
    var ctx = new Word.RequestContext();
    var ccs = ctx.document.contentControls.getByTag("Customer-Address");
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


### getByTitle(title: string)
Gets the content controls that have the specified title.

#### Syntax
```js
    contentControl.getByTitle(title: string);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`title`          | string | The title of a content control. |

#### Returns

[contentControlCollection](contentControlCollection.md)

### getItem(index: number)
Gets a content control by its index in the collection.

#### Syntax
```js
    contentControl.getItemAt(index: number);
```
#### Parameters

Parameter      | Type   | Description |
-------------- | ------ | ------------ |
`index`          | number | The index  |

#### Returns

[contentControl](contentControl.md)



[Back](#methods)







