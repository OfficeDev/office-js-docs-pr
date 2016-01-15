# Section object (JavaScript API for Word)

Represents a section in a Word document.

_Applies to: Word 2016, Word for iPad, Word for Mac_

## Properties
None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|body|[Body](body.md)|Gets the body of the section. This does not include the headerfooter and other section metadata. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getFooter(type: HeaderFooterType)](#getfootertype-headerfootertype)|[Body](body.md)|Gets one of the section's footers.|
|[getHeader(type: HeaderFooterType)](#getheadertype-headerfootertype)|[Body](body.md)|Gets one of the section's headers.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method details

### getFooter(type: HeaderFooterType)
Gets one of the section's footers.

#### Syntax
```js
sectionObject.getFooter(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|HeaderFooterType|Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'.|

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
### getHeader(type: HeaderFooterType)
Gets one of the section's headers.

#### Syntax
```js
sectionObject.getHeader(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|HeaderFooterType|Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'.|

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

## Support details

Use the [requirement set](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) in run time checks to make sure your application is supported by the host version of Word. For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 