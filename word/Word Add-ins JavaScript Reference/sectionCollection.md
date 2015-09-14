# SectionCollection
Contains a collection of [Section](section.md) objects.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|items|  array | Gets an array of section objects. |


## Relationships
None  

## Methods

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)| [Section](section.md)  | Gets a section object by its index in the collection. |
|[load(param: option)](#loadparam-option)|void|Fills the section collection proxy object created in the JavaScript layer with property and object values specified in the parameter.|



## API Specification

### getItem(index: number)

Gets a section object by its index in the collection.

#### Syntax
```js
    sectionCollection.getItem(index);
```
#### Parameters

| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|  A number that identifies the index location of a section object.  |

#### Returns

[Paragraph](paragraph.md)  


#### Example
```js
    var ctx = new Word.RequestContext();

    var mySections  = ctx.document.sections;
    ctx.load(mySections);

    var myFooter = mySections.getItem(0).getFooter("primary");
    myFooter.insertText("this is a footer!!","end");

    ctx.executeAsync().then(
        function(){
           console.log("Success!!");
        }
    );
```

[Back](#methods)

### load(param: option)
Fills the section collection proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    sectionCollection.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)