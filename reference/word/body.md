# Body object (JavaScript API for Word)

Represents the body of a document or a section.

_Applies to: Word 2016, Word for iPad, Word for Mac_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|style|string|Gets or sets the style used for the body. This is the name of the pre-installed or custom style.|
|text|string|Gets the text of the body. Use the insertText method to insert text. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of rich text content control objects that are in the body. Read-only.|
|font|[Font](font.md)|Gets the text format of the body. Use this to get and set font name, size, color, and other properties. Read-only.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects that are in the body. The collection does not include floating images. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of paragraph objects that are in the body. Read-only.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the body. Returns null if there isn't a parent content control. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the body object. The user can perform the undo operation on the cleared content.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the body object.|
|[getOoxml()](#getooxml)|string|Gets the OOXML (Office Open XML) representation of the body object.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location. A break can only be inserted into the main document body, except if it is a line break which can be inserted into any body object. The insertLocation value can be 'Start' or 'End'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the body object with a Rich Text content control.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML or wordProcessingML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the body object. The search results are a collection of range objects.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects the body and navigates the Word UI to it. The selectionMode values can be 'Select', 'Start', or 'End'.|

## Method details

### clear()
Clears the contents of the body object. The user can perform the undo operation on the cleared content.

#### Syntax
```js
bodyObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to clear the contents of the body.
    body.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the body contents.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

The [Silly stories](https://aka.ms/sillystorywordaddin) add-in sample shows how the **clear** method can be used to clear the contents of a document.

### getHtml()
Gets the HTML representation of the body object.

#### Syntax
```js
bodyObject.getHtml();
```

#### Parameters
None

#### Returns
string

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body HTML contents: " + bodyHTML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
Gets the OOXML (Office Open XML) representation of the body object.

#### Syntax
```js
bodyObject.getOoxml();
```

#### Parameters
None

#### Returns
string

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body OOXML contents: " + bodyOOXML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location. A break can only be inserted into the main document body, except if it is a line break which can be inserted into any body object. The insertLocation value can be 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Required. The break type to add to the body.|
|insertLocation|InsertLocation|Required. The value can be 'Start' or 'End'.|

#### Returns
void

#### Additional details
With the exception of line breaks, you can not insert a break in headers, footers, footnotes, endnotes, comments, and textboxes.

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (ctx) {

    // Create a proxy object for the document body.
    var body = ctx.document.body;

    // Queue a commmand to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log('Added a page break at the start of the document body.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### insertContentControl()
Wraps the body object with a Rich Text content control.

#### Syntax
```js
bodyObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to wrap the body in a content control.
    body.insertContentControl();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped the body in a content control.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded file contents to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert base64 encoded .docx at the beginning of the content body.
    // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.
    body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

The [Silly stories](https://aka.ms/sillystorywordaddin) add-in sample shows how the **insertFileFromBase64** method can be used to insert docx files from a service.

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted in the document.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert HTML in to the beginning of the body.
    body.insertHtml('<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts a picture into the body at the specified location. The insertLocation value can be 'Start' or 'End'.

#### Syntax
bodyObject.insertInlinePictureFromBase64(image, insertLocation);

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted in the body.|
|insertLocation|InsertLocation|Required. The value can be 'Start' or 'End'.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML or wordProcessingML at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML or wordProcessingML to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert OOXML in to the beginning of the body.
    body.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the document body.');
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
Read [Create better add-ins for Word with Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) for guidance on working with OOXML. The [Word-Add-in-DocumentAssembly][body.insertOoxml] sample shows how you can use this API to assemble a document.

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Start' or 'End'.|

#### Returns
[Paragraph](paragraph.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added at the end of the document body.');
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
The [Word-Add-in-DocumentAssembly][body.insertParagraph] sample shows how you can use the insertParagraph method to assemble a document.

### insertText(text: string, insertLocation: InsertLocation)
Inserts text into the body at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
bodyObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. Text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the beginning of the document body.');
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

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Performs a search with the specified search options on the scope of the body object. The search results are a collection of range objects.

#### Syntax
```js
bodyObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', {matchCase: false});

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length +
                      '; we highlighted the results.';

        // Queue a command to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';
          searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log(results);
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

#### Additional information
The [Word-Add-in-DocumentAssembly][body.search] sample provides another example of how to search a document.

### select(selectionMode: SelectionMode)
Selects the body and navigates the Word UI to it. The selectionMode values can be 'Select', 'Start', or 'End'.

#### Syntax
```js
bodyObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to select the document body. The Word UI will
    // move to the selected document body.
    body.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Property access examples

### Get the text property on the body object
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load the text in document body.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### Get the style and the font size, font name, and font color properties on the body object.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
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


[body.insertOoxml]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L127 "insert OOXML"
[body.insertParagraph]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L153 "insert paragraph"
[body.search]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L261 "body search"