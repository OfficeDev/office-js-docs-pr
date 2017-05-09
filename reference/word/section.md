# Section Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a section in a Word document.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|body|[Body](body.md)|Gets the body object of the section. This does not include the headerfooter and other section metadata. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getFooter(type: string)](#getfootertype-string)|[Body](body.md)|Gets one of the section's footers.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[getHeader(type: string)](#getheadertype-string)|[Body](body.md)|Gets one of the section's headers.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[getNext()](#getnext)|[Section](section.md)|Gets the next section. Throws if this section is the last one.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getNextOrNullObject()](#getnextornullobject)|[Section](section.md)|Gets the next section. Returns a null object if this section is the last one.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### getFooter(type: string)
Gets one of the section's footers.

#### Syntax
```js
sectionObject.getFooter(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|type|string|Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'. Possible values are: `Primary` Returns the header or footer on all pages of a section, with the first page or odd pages excluded if they are different.,`FirstPage` Returns the header or footer on the first page of a section.,`EvenPages` Returns all headers or footers on even-numbered pages of a section.|

#### Returns
[Body](body.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
	
	// Create a proxy sectionsCollection object.
	var mySections = context.document.sections;
	
	// Queue a commmand to load the sections.
	context.load(mySections, 'body/style');
	
	// Synchronize the document state by executing the queued commands, 
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		
		// Create a proxy object the primary footer of the first section. 
		// Note that the footer is a body object.
		var myFooter = mySections.items[0].getFooter("primary");
		
		// Queue a command to insert text at the end of the footer.
		myFooter.insertText("This is a footer.", Word.InsertLocation.end);
		
		// Queue a command to wrap the header in a content control.
		myFooter.insertContentControl();
							  
		// Synchronize the document state by executing the queued commands, 
		// and return a promise to indicate task completion.
		return context.sync().then(function () {
			console.log("Added a footer to the first section.");
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

### getHeader(type: string)
Gets one of the section's headers.

#### Syntax
```js
sectionObject.getHeader(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|type|string|Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'. Possible values are: `Primary` Returns the header or footer on all pages of a section, with the first page or odd pages excluded if they are different.,`FirstPage` Returns the header or footer on the first page of a section.,`EvenPages` Returns all headers or footers on even-numbered pages of a section.|

#### Returns
[Body](body.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary header of the first section. 
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");
        
        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
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

### getNext()
Gets the next section. Throws if this section is the last one.

#### Syntax
```js
sectionObject.getNext();
```

#### Parameters
None

#### Returns
[Section](section.md)

### getNextOrNullObject()
Gets the next section. Returns a null object if this section is the last one.

#### Syntax
```js
sectionObject.getNextOrNullObject();
```

#### Parameters
None

#### Returns
[Section](section.md)

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
