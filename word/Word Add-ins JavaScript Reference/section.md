# Section 
Represents a section in a Word document.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|body|  [Body](body.md)   | Gets the body of the section. This does not include the header/footer and other section metadata.  |

## Relationships
None

## Methods

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[getFooter(type: string)](#getfootertype-string)| [Body](body.md) | Gets the section footer.|     
|[getHeader(type: string)](#getheadertype-header)| [Body](body.md) |Gets the section header. |
|[load(param: option)](#loadparam-option)|void|Fills the section proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## API Specification


### getFooter(type: string) 

Gets the section footer.

#### Syntax
```js
    section.getFooter(type);
```
#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`type`         | string | Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'. |

#### Returns

[Body](body.md)


#### Example

```js
    //Insert text in to the footer of the first section.

    var ctx = new Word.RequestContext();

    var mySections  = ctx.document.sections;
    ctx.load(mySections);

    var myFooter = mySections.getItem(0).getFooter("primary");
    myFooter.insertText("this is a footer","end");

    ctx.executeAsync().then(
        function(){
            console.log("Success!!");
        }
    );

```
[Back](#methods)


### getHeader(type: string) 

Gets the section header.

#### Syntax
```js
    section.getHeader(type);
```

#### Parameters

Parameter      | Type   | Description|
-------------- | ------ | ------------|
`type`         | string | Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'. |


#### Returns

[Body](body.md).


#### Example

```js
    //Insert text in to the footer of the first section.

    var ctx = new Word.RequestContext();

    var mySections  = ctx.document.sections;
    ctx.load(mySections);

    var myHeader = mySections.getItem(0).getHeader("primary");
    myHeader.insertText("this is a header!!","end");

    ctx.executeAsync().then(
        function(){
            console.log("Success!!");
        }
    );
```
[Back](#methods)

### load(param: option)
Fills the section proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    section.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)