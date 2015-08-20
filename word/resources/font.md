# Font

Represents a font.

## Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|bold| bool  | Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.|
|color| string  | Gets or sets the color for the specified font. You can provide the value as either the hexidecimal color value or the color name. |
|doubleStrikeThrough| bool | Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.| 
|highlightColor| string  | Gets or sets the highlight color for the specified font. You can provide the value as either the hexidecimal color value or the color name. |
|italic| bool  | Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false. |
|name| string  | Gets or sets a value that represents the name of the font. |
|size| number  | Gets or sets a value that represents the font size in points.|
|strikeThrough| bool  | Gets or sets a value that indicates whether the font has a strike through. True if the font is formatted as strikethrough text, otherwise, false. |
|subscript| bool  |Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false. |
|superscript| bool  | Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false. |
|underline|  bool  | Gets or sets a value that indicates whether the font is underlined. True if the font is underlined, otherwise, false. |

## Methods

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[load(param: option)](#loadparam-option)|void|Fills the font proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## API Specification

### load(param: option)

Fills the font proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    font.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)


### Getter and Setter Examples

#### Change font properties
```js
    // insert a paragraph and use the font object to change font properties

    var ctx = new Word.RequestContext();

    var myPar = ctx.document.body.insertParagraph("Here is some text!","end");
    myPar.font.bold = true;
    myPar.font.italic = true;
    myPar.font.color = "#00FF00";  // lime green!
    myPar.font.doubleStrikeThrough = true;

    ctx.executeAsync().then(
         function () {
             console.log("Success!!");
         },
         function (result) {
             console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
         }
    );

```



