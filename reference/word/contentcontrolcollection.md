# ContentControlCollection object (JavaScript API for Word)

Contains a collection of ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.

_Applies to: Word 2016, Word for iPad, Word for Mac_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[ContentControl[]](contentcontrol.md)|A collection of contentControl objects. Read-only.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getById(id: number)](#getbyidid-number)|[ContentControl](contentcontrol.md)|Gets a content control by its identifier.|
|[getByTag(tag: string)](#getbytagtag-string)|[ContentControlCollection](contentcontrolcollection.md)|Gets the content controls that have the specified tag.|
|[getByTitle(title: string)](#getbytitletitle-string)|[ContentControlCollection](contentcontrolcollection.md)|Gets the content controls that have the specified title.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method details

### getById(id: number)
Gets a content control by its identifier.

#### Syntax
```js
contentControlCollectionObject.getById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|number|Required. A content control identifier.|

#### Returns
[ContentControl](contentcontrol.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

	// Create a proxy object for the content control that contains a specific id.
	var contentControl = context.document.contentControls.getById(30086310);

	// Queue a command to load the text property for a content control.
	context.load(contentControl, 'text');

	// Synchronize the document state by executing the queued commands,
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		console.log('The content control with that Id has been found in this document.');
	});
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
	}
});
```

### getByTag(tag: string)
Gets the content controls that have the specified tag.

#### Syntax
```js
contentControlCollectionObject.getByTag(tag);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|tag|string|Required. A tag set on a content control.|

#### Returns
[ContentControlCollection](contentcontrolcollection.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');

    // Queue a command to load the text property for all of content controls with a specific tag.
    context.load(contentControlsWithTag, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log("There isn't a content control with a tag of Customer-Address in this document.");
        } else {
            console.log('The first content control with the tag of Customer-Address has this text: ' + contentControlsWithTag.items[0].text);
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

#### Additional information
The [Word-Add-in-DocumentAssembly][contentControls.getByTag] sample has another example of using the getByTag method.


### getByTitle(title: string)
Gets the content controls that have the specified title.

#### Syntax
```js
contentControlCollectionObject.getByTitle(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|string|Required. The title of a content control.|

#### Returns
[ContentControlCollection](contentcontrolcollection.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific title.
    var contentControlsWithTitle = context.document.contentControls.getByTitle('Enter Customer Address Here');

    // Queue a command to load the text property for all of content controls with a specific title.
    context.load(contentControlsWithTitle, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTitle.items.length === 0) {
            console.log("There isn't a content control with a title of 'Enter Customer Address Here' in this document.");
        } else {
            console.log("The first content control with the title of 'Enter Customer Address Here' has this text: " + contentControlsWithTitle.items[0].text);
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

#### Additional information
The [Word-Add-in-DocumentAssembly][contentControls.getByTitle] sample has another example of using the getByTitle method.

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

    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    // Queue a command to load the id property for all of the content controls.
    context.load(contentControls, 'id');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control.
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' +
                        '   ----- appearance: ' + contentControls.items[0].appearance +
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
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

The [Silly stories](https://aka.ms/sillystorywordaddin) add-in sample shows how the **load** method is used to load the content control collection with the **tag** and **title** properties.

## Support details

Use the [requirement set](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) in run time checks to make sure your application is supported by the host version of Word. For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx).


[contentControls.getByTag]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L300 "get by tag"
[contentControls.getByTitle]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L331 "get by title"

