# Range Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a contiguous area in a document.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|hyperlink|string|Gets the first hyperlink in the range, or sets a hyperlink on the range. Existing hyperlinks in this range are deleted when you set a new hyperlink.|
|isEmpty|bool|Checks whether the range length is zero. Read-only.|
|style|string|Gets or sets the style used for the range. This is the name of the pre-installed or custom style.|
|text|string|Gets the text of the range. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the range. Read-only.|
|font|[Font](font.md)|Gets the text format of the range. Use this to get and set font name, size, color, and other properties. Read-only.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inline picture objects in the range. Read-only.|
|lists|[ListCollection](listcollection.md)|Gets the collection of list objects in the range. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of paragraph objects in the range. Read-only.|
|parentBody|[Body](body.md)|Gets the parent body of the range. Read-only.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the range. Returns null if there isn't a parent content control. Read-only.|
|parentTable|[Table](table.md)|Gets the table that contains the range. Returns null if it is not contained in a table. Read-only.|
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the range. Returns null if it is not contained in a table cell. Read-only.|
|tables|[TableCollection](tablecollection.md)|Gets the collection of table objects in the range. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the range object. The user can perform the undo operation on the cleared content.|
|[compareLocationWith(range: Range)](#comparelocationwithrange-range)|[LocationRelation](locationrelation.md)|Compares this range's location with another range's location.|
|[delete()](#delete)|void|Deletes the range and its content from the document.|
|[expandTo(range: Range)](#expandtorange-range)|void|Expands the range in either direction to cover another range.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the range object.|
|[getHyperlinkRanges()](#gethyperlinkranges)|[RangeCollection](rangecollection.md)|Gets hyperlink child ranges within the range.|
|[getNextTextRange(punctuationMarks: string[], trimSpacing: bool)](#getnexttextrangepunctuationmarks-string-trimspacing-bool)|[Range](range.md)|Gets the next text range by using punctuation marks andor space character.|
|[getOoxml()](#getooxml)|string|Gets the OOXML representation of the range object.|
|[getRange(rangeLocation: string)](#getrangerangelocation-string)|[Range](range.md)|Clones the range, or gets the starting or ending point of the range as a new range.|
|[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](#gettextrangespunctuationmarks-string-trimspacing-bool)|[RangeCollection](rangecollection.md)|Gets the text child ranges in the range by using punctuation marks andor space character.|
|[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocation-string)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Replace', 'Before' or 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the range object with a rich text content control.|
|[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-insertlocation-string)|[Range](range.md)|Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)|[Range](range.md)|Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-string)|[InlinePicture](inlinepicture.md)|Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-string)|[Range](range.md)|Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])](#inserttablerowcount-number-columncount-number-insertlocation-string-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|
|[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-string)|[Range](range.md)|Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|
|[intersectWith(range: Range)](#intersectwithrange-range)|void|Shrinks the range to the intersection of the range with another range.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.|
|[select(selectionMode: string)](#selectselectionmode-string)|void|Selects and navigates the Word UI to the range.|
|[split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)](#splitdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimspacing-bool)|[RangeCollection](rangecollection.md)|Splits the range into child ranges by using delimiters.|

## Method Details


### clear()
Clears the contents of the range object. The user can perform the undo operation on the cleared content.

#### Syntax
```js
rangeObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to clear the contents of the proxy range object.
    range.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the selection (range object)');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### compareLocationWith(range: Range)
Compares this range's location with another range's location.

#### Syntax
```js
rangeObject.compareLocationWith(range);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|range|Range|Required. The range to compare with this range.|

#### Returns
[LocationRelation](locationrelation.md)

### delete()
Deletes the range and its content from the document.

#### Syntax
```js
rangeObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to delete the range object.
    range.delete();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Deleted the selection (range object)');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### expandTo(range: Range)
Expands the range in either direction to cover another range.

#### Syntax
```js
rangeObject.expandTo(range);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|range|Range|Required. Another range.|

#### Returns
void

### getHtml()
Gets the HTML representation of the range object.

#### Syntax
```js
rangeObject.getHtml();
```

#### Parameters
None

#### Returns
string

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to get the HTML of the current selection.
    var html = range.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The HTML read from the document was: ' + html.value);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### getHyperlinkRanges()
Gets hyperlink child ranges within the range.

#### Syntax
```js
rangeObject.getHyperlinkRanges();
```

#### Parameters
None

#### Returns
[RangeCollection](rangecollection.md)

### getNextTextRange(punctuationMarks: string[], trimSpacing: bool)
Gets the next text range by using punctuation marks andor space character.

#### Syntax
```js
rangeObject.getNextTextRange(punctuationMarks, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|punctuationMarks|string[]|Required. The punctuation marks and/or space character as an array of strings.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included.|

#### Returns
[Range](range.md)

### getOoxml()
Gets the OOXML representation of the range object.

#### Syntax
```js
rangeObject.getOoxml();
```

#### Parameters
None

#### Returns
string

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to get the OOXML of the current selection.
    var ooxml = range.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The OOXML read from the document was:  ' + ooxml.value);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### getRange(rangeLocation: string)
Clones the range, or gets the starting or ending point of the range as a new range.

#### Syntax
```js
rangeObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|string|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.  Possible values are: Whole, Start, End|

#### Returns
[Range](range.md)

### getTextRanges(punctuationMarks: string[], trimSpacing: bool)
Gets the text child ranges in the range by using punctuation marks andor space character.

#### Syntax
```js
rangeObject.getTextRanges(punctuationMarks, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|punctuationMarks|string[]|Required. The punctuation marks and/or space character as an array of strings.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)

### insertBreak(breakType: string, insertLocation: string)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Replace', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|string|Required. The break type to add. Possible values are: `Page` Page break at the insertion point.,`Column` Column break at the insertion point.,`Next` Section break on next page.,`SectionContinuous` New section without a corresponding page break.,`SectionEven` Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.,`SectionOdd` Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.,`Line` Line break.,`LineClearLeft` Line break.,`LineClearRight` Line break.,`TextWrapping` Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.|
|insertLocation|string|Required. The value can be 'Replace', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a page break after the selected text.
    range.insertBreak('page', 'After');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted a page break after the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### insertContentControl()
Wraps the range object with a rich text content control.

#### Syntax
```js
rangeObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert a content control around the selected text,
    // and create a proxy content control object. We'll update the properties
    // on the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = "Customer-Address";
    myContentControl.title = "Enter Customer Address Here:";
    myContentControl.style = "Normal";
    myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
    myContentControl.cannotEdit = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped a content control around the selected text.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### insertFileFromBase64(base64File: string, insertLocation: string)
Inserts a document at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|string|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### insertHtml(html: string, insertLocation: string)
Inserts HTML at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted.|
|insertLocation|string|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)
Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted.|
|insertLocation|string|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: string)
Inserts OOXML at the specified location.  The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML to be inserted.|
|insertLocation|string|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

*Additional information* 
Read [Create better add-ins for Word with Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) for guidance on working with OOXML.


### insertParagraph(paragraphText: string, insertLocation: string)
Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Paragraph](paragraph.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added to the end of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])
Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertTable(rowCount, columnCount, insertLocation, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowCount|number|Required. The number of rows in the table.|
|columnCount|number|Required. The number of columns in the table.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[Table](table.md)

### insertText(text: string, insertLocation: string)
Inserts text at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
rangeObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. Text to be inserted.|
|insertLocation|string|Required. The value can be 'Replace', 'Start', 'End', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the end of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### intersectWith(range: Range)
Shrinks the range to the intersection of the range with another range.

#### Syntax
```js
rangeObject.intersectWith(range);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|range|Range|Required. Another range.|

#### Returns
void

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

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Performs a search with the specified searchOptions on the scope of the range object. The search results are a collection of range objects.

#### Syntax
```js
rangeObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Queue a command to select the HTML that was inserted.
    range.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### select(selectionMode: string)
Selects and navigates the Word UI to the range.

#### Syntax
```js
rangeObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|string|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.  Possible values are: Select, Start, End|

#### Returns
void

### split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)
Splits the range into child ranges by using delimiters.

#### Syntax
```js
rangeObject.split(delimiters, multiParagraphs, trimDelimiters, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|delimiters|string[]|Required. The delimiters as an array of strings.|
|multiParagraphs|bool|Optional. Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.|
|trimDelimiters|bool|Optional. Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)
