# ContentControl Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|appearance|string|Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'. Possible values are: `BoundingBox` Represents a content control shown as a shaded rectangle or bounding box (with optional title).,`Tags` Represents a content control shown as start and end markers.,`Hidden` Represents a content control that is not shown.|
|cannotDelete|bool|Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.|
|cannotEdit|bool|Gets or sets a value that indicates whether the user can edit the contents of the content control.|
|color|string|Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.|
|placeholderText|string|Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.|
|removeWhenEdited|bool|Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.|
|style|string|Gets or sets the style used for the content control. This is the name of the pre-installed or custom style.|
|subtype|string|Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only. Possible values are: RichText, Unknown, RichTextInline, RichTextParagraphs, RichTextTableCell, RichTextTableRow, RichTextTable, PlainTextInline, PlainTextParagraph, Picture, BuildingBlockGallery, CheckBox, ComboBox, DropDownList, DatePicker, RepeatingSection, PlainText.|
|tag|string|Gets or sets a tag to identify a content control.|
|text|string|Gets the text of the content control. Read-only.|
|title|string|Gets or sets the title for a content control.|
|type|string|Gets the content control type. Only rich text content controls are supported currently. Read-only. Possible values are: RichText, Unknown, RichTextInline, RichTextParagraphs, RichTextTableCell, RichTextTableRow, RichTextTable, PlainTextInline, PlainTextParagraph, Picture, BuildingBlockGallery, CheckBox, ComboBox, DropDownList, DatePicker, RepeatingSection, PlainText.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the content control. Read-only.|
|font|[Font](font.md)|Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.|
|id|[uint](uint.md)|Gets an integer that represents the content control identifier. Read-only.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.|
|lists|[ListCollection](listcollection.md)|Gets the collection of list objects in the content control. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Get the collection of paragraph objects in the content control. Read-only.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the content control. Returns null if there isn't a parent content control. Read-only.|
|parentTable|[Table](table.md)|Gets the table that contains the content control. Returns null if it is not contained in a table. Read-only.|
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the content control. Returns null if it is not contained in a table cell. Read-only.|
|tables|[TableCollection](tablecollection.md)|Gets the collection of table objects in the content control. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clears the contents of the content control. The user can perform the undo operation on the cleared content.|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|void|Deletes the content control and its content. If keepContent is set to true, the content is not deleted.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the content control object.|
|[getOoxml()](#getooxml)|string|Gets the Office Open XML (OOXML) representation of the content control object.|
|[getRange(rangeLocation: string)](#getrangerangelocation-string)|[Range](range.md)|Gets the whole content control, or the starting or ending point of the content control, as a range.|
|[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](#gettextrangespunctuationmarks-string-trimspacing-bool)|[RangeCollection](rangecollection.md)|Gets the text ranges in the content control by using punctuation marks andor space character.|
|[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocation-string)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.|
|[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-insertlocation-string)|[Range](range.md)|Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)|[Range](range.md)|Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-string)|[InlinePicture](inlinepicture.md)|Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-string)|[Range](range.md)|Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|
|[insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])](#inserttablerowcount-number-columncount-number-insertlocation-string-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|
|[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-string)|[Range](range.md)|Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.|
|[select(selectionMode: string)](#selectselectionmode-string)|void|Selects the content control. This causes Word to scroll to the selection.|
|[split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)](#splitdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimspacing-bool)|[RangeCollection](rangecollection.md)|Splits the content control into child ranges by using delimiters.|

## Method Details


### clear()
Clears the contents of the content control. The user can perform the undo operation on the cleared content.

#### Syntax
```js
contentControlObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples

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


### delete(keepContent: bool)
Deletes the content control and its content. If keepContent is set to true, the content is not deleted.

#### Syntax
```js
contentControlObject.delete(keepContent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|keepContent|bool|Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.|

#### Returns
void

#### Examples

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
            
            // Queue a command to delete the first content control. The
            // contents will remain in the document.
            contentControls.items[0].delete(true);
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



### getHtml()
Gets the HTML representation of the content control object.

#### Syntax
```js
contentControlObject.getHtml();
```

#### Parameters
None

#### Returns
string

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
    
    // Queue a command to load the tag property for all of content controls. 
    context.load(contentControlsWithTag, 'tag');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the HTML contents of the first content control.
            var html = contentControlsWithTag.items[0].getHtml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control HTML: ' + html.value);
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


### getOoxml()
Gets the Office Open XML (OOXML) representation of the content control object.

#### Syntax
```js
contentControlObject.getOoxml();
```

#### Parameters
None

#### Returns
string

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
            // Queue a command to get the OOXML contents of the first content control.
            var ooxml = contentControls.items[0].getOoxml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control OOXML: ' + ooxml.value);
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


### getRange(rangeLocation: string)
Gets the whole content control, or the starting or ending point of the content control, as a range.

#### Syntax
```js
contentControlObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|string|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.  Possible values are: Whole, Start, End|

#### Returns
[Range](range.md)

### getTextRanges(punctuationMarks: string[], trimSpacing: bool)
Gets the text ranges in the content control by using punctuation marks andor space character.

#### Syntax
```js
contentControlObject.getTextRanges(punctuationMarks, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|punctuationMarks|string[]|Required. The punctuation marks and/or space character as an array of strings.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)

### insertBreak(breakType: string, insertLocation: string)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.

#### Syntax
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|string|Required. Type of break. Possible values are: `Page` Page break at the insertion point.,`Column` Column break at the insertion point.,`Next` Section break on next page.,`SectionContinuous` New section without a corresponding page break.,`SectionEven` Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.,`SectionOdd` Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.,`Line` Line break.,`LineClearLeft` Line break.,`LineClearRight` Line break.,`TextWrapping` Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.|
|insertLocation|string|Required. The value can be 'Start', 'End', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a commmand to load the id property for all of content controls. 
    context.load(contentControls, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion. We now will have 
    // access to the content control collection.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a page break after the first content control. 
            contentControls.items[0].insertBreak('page', "After");
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion. 
            return context.sync()
                .then(function () {
                    console.log('Inserted a page break after the first content control.');    
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


### insertFileFromBase64(base64File: string, insertLocation: string)
Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

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
            // Queue a command to put HTML into the contents of the first content control.
            contentControls.items[0].insertHtml('<strong>HTML content inserted into the content control.</strong>', 'Start');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted HTML in the first content control.');
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


### insertHtml(html: string, insertLocation: string)
Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted in to the content control.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)
Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted in the content control.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[InlinePicture](inlinepicture.md)

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
            // Queue a command to put OOXML into the contents of the first content control.
            contentControls.items[0].insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", "End");
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted OOXML in the first content control.');
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

*Additional information*
Read [Create better add-ins for Word with Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) for guidance on working with OOXML.


### insertOoxml(ooxml: string, insertLocation: string)
Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML to be inserted in to the content control.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: string)
Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragrph text to be inserted.|
|insertLocation|string|Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Paragraph](paragraph.md)

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
            // Queue a command to insert a paragraph after the first content control. 
            contentControls.items[0].insertParagraph('Text of the inserted paragraph.', 'After');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted a paragraph after the first content control.');
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


### insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])
Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
contentControlObject.insertTable(rowCount, columnCount, insertLocation, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rowCount|number|Required. The number of rows in the table.|
|columnCount|number|Required. The number of columns in the table.|
|insertLocation|string|Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[Table](table.md)

### insertText(text: string, insertLocation: string)
Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. The text to be inserted in to the content control.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

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
            // Queue a command to replace text in the first content control. 
            contentControls.items[0].insertText('Replaced text in the first content control.', 'Replace');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Replaced text in the first content control.');
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

The [Silly stories](https://aka.ms/sillystorywordaddin) add-in sample shows how to use the **insertText** method.


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
    
    // Create a proxy range object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to create the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = 'Customer-Address';
    myContentControl.title = ' has t';
    myContentControl.style = 'Heading 2';
    myContentControl.insertText('One Microsoft Way, Redmond, WA 98052', 'replace');
    myContentControl.cannotEdit = true;
    
    // Queue a command to load the id property for the content control you created.
    context.load(myContentControl, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Created content control with id: ' + myContentControl.id);
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
Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.

#### Syntax
```js
contentControlObject.search(searchText, searchOptions);
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
            // Queue a command to select the first content control.
            contentControls.items[0].select();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Selected the first content control.');
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


### select(selectionMode: string)
Selects the content control. This causes Word to scroll to the selection.

#### Syntax
```js
contentControlObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|string|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.  Possible values are: Select, Start, End|

#### Returns
void

### split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)
Splits the content control into child ranges by using delimiters.

#### Syntax
```js
contentControlObject.split(delimiters, multiParagraphs, trimDelimiters, trimSpacing);
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
### Property access examples

*Load all of the content control properties*

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
