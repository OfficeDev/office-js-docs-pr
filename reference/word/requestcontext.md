# RequestContext object (JavaScript API for Word)

The RequestContext object facilitates requests to the Word application from the Word add-in since the two applications run in different processes. 

_Applies to: Word 2016, Word for iPad, Word for Mac_

## Properties
None

## Methods

| Method         | Return Type    |Description|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Fills the proxy object created in JavaScript layer with property and options specified in the parameter.|
|[sync()](#sync)  |Promise object |Submits the request queue to Word and returns a promise object, which can be used for chaining further actions.|

## Method details

### load(object: object, option: object)
Fills the proxy object created in JavaScript layer with property and options specified in the parameter.

#### Syntax
```js
requestContextObject.load(object, loadOption);
```

#### Parameters
| Parameter       | Type    |Description|
|:----------------|:--------|:----------|
|object|object|Optional. Specify the name of the object to be loaded.|
|option|[loadOption](loadoption.md)|Optional, but is used as best practice. Specify the load options such as select, expand, skip and top. |

#### Returns
void

##### Examples

The following example shows how the request context is used to load the text property on a paragraph collection.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
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

### sync() 
Submits the request queue to Word and returns a promise object, which can be used for chaining further actions.

#### Syntax
```js
requestContextObject.sync();
```

#### Parameters
None

#### Returns
Promise object.

#### Examples

The following example shows the sync method used twice: 1) load the content controls collection with the text property for each content control, and 2) clearing the contents of the first content control in the collection.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
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

## Support details

Use the [requirement set](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) in run time checks to make sure your application is supported by the host version of Word. For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 