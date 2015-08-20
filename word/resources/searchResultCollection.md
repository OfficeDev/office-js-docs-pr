# SearchResultCollection

Contains a collection of [Range](range.md) objects as a result of a search operation.


## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|items|  array | Gets an array of range objects. |


## Relationships
None  

## Methods

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)| [Range](range.md)   | Gets a range object by its index in the collection. |
|[load(param: option)](#loadparam-option)|void|Fills the search result collection proxy object created in the JavaScript layer with property and object values specified in the parameter.|


## API Specification

### getItem(index: number)

Gets a range object by its index in the collection.

#### Syntax
```js
    searchResultCollection.getItem(index);
```
#### Parameters

| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|  A number that identifies the index location of a range object.  |

#### Returns

[Range](range.md)


[Back](#methods)


### load(param: option)
Fills the search collection proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    searchResultCollection.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)





#### Example
```js

    ///Search example, returns a collection of ranges

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



