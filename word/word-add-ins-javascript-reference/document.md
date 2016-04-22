# Document Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|saved|bool|Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|body|[Body](body.md)|Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the current document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.|
|sections|[SectionCollection](sectioncollection.md)|Gets the collection of section objects in the document. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getSelection()](#getselection)|[Range](range.md)|Gets the current selection of the document. Multiple selections are not supported.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[open()](#open)|void|Open the document.|
|[save()](#save)|void|Saves the document. This will use the Word default file naming convention if the document has not been saved before.|

## Method Details


### getSelection()
Gets the current selection of the document. Multiple selections are not supported.

#### Syntax
```js
documentObject.getSelection();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    var textSample = 'This is an example of the insert text method. This is a method ' + 
        'which allows users to insert text into a selection. It can insert text into a ' +
        'relative location or it can overwrite the current selection. Since the ' +
        'getSelection method returns a range object, look up the range object documentation ' +
        'for everything you can do with a selection.';
    
    // Create a range proxy object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert text at the end of the selection.
    range.insertText(textSample, Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted the text at the end of the selection.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;
    
    // Queue a command to load content control properties.
    context.load(thisDocument, 'contentControls/id, contentControls/text, contentControls/tag');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (thisDocument.contentControls.items.length !== 0) {
            for (var i = 0; i < thisDocument.contentControls.items.length; i++) {
                console.log(thisDocument.contentControls.items[i].id);
                console.log(thisDocument.contentControls.items[i].text);
                console.log(thisDocument.contentControls.items[i].tag);
            }
        } else {
            console.log('No content controls in this document.');
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### open()
Open the document.

#### Syntax
```js
documentObject.open();
```

#### Parameters
None

#### Returns
void

### save()
Saves the document. This will use the Word default file naming convention if the document has not been saved before.

#### Syntax
```js
documentObject.save();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;

    // Queue a commmand to load the document save state (on the saved property).
    context.load(thisDocument, 'saved');    
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (thisDocument.saved === false) {
            // Queue a command to save this document.
            thisDocument.save();
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Saved the document');
            });
        } else {
            console.log('The document has not changed since the last save.');
        }
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
