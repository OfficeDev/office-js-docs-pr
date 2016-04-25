# Font Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a font.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.|WordApi1.1||
|color|string|Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.|WordApi1.1||
|doubleStrikeThrough|bool|Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.|WordApi1.1||
|highlightColor|string|Gets or sets the highlight color for the specified font. You can provide the value as either in the '#RRGGBB' format or the color name.|WordApi1.1||
|italic|bool|Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.|WordApi1.1||
|name|string|Gets or sets a value that represents the name of the font.|WordApi1.1||
|strikeThrough|bool|Gets or sets a value that indicates whether the font has a strike through. True if the font is formatted as strikethrough text, otherwise, false.|WordApi1.1||
|subscript|bool|Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.|WordApi1.1||
|superscript|bool|Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.|WordApi1.1||
|underline|string|Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined. Possible values are: `None` No underline.,`Single` A single underline. This is the default value.,`Word` Only underline individual words.,`Double` A double underline.,`Dotted` A dotted underline.,`Hidden` A hidden underline.,`Thick` A single thick underline.,`DashLine` A single dash underline.,`DotLine` A single dot underline.,`DotDashLine` An alternating dot-dash underline.,`TwoDotDashLine` An alternating dot-dot-dash underline.,`Wave` A single wavy underline.|WordApi1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|size|[float](float.md)|Gets or sets a value that represents the font size in points.|WordApi1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|WordApi1.1|

## Method Details


### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the font property for all of the paragraphs.
    context.load(paragraphs, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Create a proxy object for the font object on the first paragraph in the collection.
        var font = paragraphs.items[0].font;

        // Queue a set of property value changes on the font proxy object.
        font.size = 32;
        font.bold = true;
        font.color = '#0000ff';
        font.highlightColor = '#ffff00';

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('The font has changed.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Property access examples

*Change the font name*

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the current selection's font name.
    selection.font.name = 'Arial';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font name has changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Change the font color* 

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the font color of the current selection.
    selection.font.color = 'blue';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font color of the selection has been changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Change the font size*


```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to change the current selection's font size.
    selection.font.size = 20;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font size has changed.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Highlight selected text*

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to highlight the current selection.
    selection.font.highlightColor = '#FFFF00'; // Yellow

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection has been highlighted.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Bold format text*

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to make the current selection bold.
    selection.font.bold = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection is now bold.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

*Underline format text*

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to underline the current selection.
    selection.font.underline = Word.UnderlineType.single;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has an underline style.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Strike format text*

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to strikethrough the font of the current selection.
    selection.font.strikeThrough = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has a strikethrough.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
