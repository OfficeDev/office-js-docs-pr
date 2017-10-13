# Paragraph Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a single paragraph in a selection, range, content control, or document body.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[alignment](enums.md)|string|Gets or sets the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'. Possible values are: `Unknown` Unknown alignment.,`Left` Alignment to the left.,`Centered` Alignment to the center.,`Right` Alignment to the right.,`Justified` Fully justified alignment.|[1.1][]|
|firstLineIndent|float|Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.|[1.1][]|
|isLastParagraph|bool|Indicates the paragraph is the last one inside its parent body. Read-only.|[1.3][]|
|isListItem|bool|Checks whether the paragraph is a list item. Read-only.|[1.3][]|
|leftIndent|float|Gets or sets the left indent value, in points, for the paragraph.|[1.1][]|
|lineSpacing|float|Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.|[1.1][]|
|lineUnitAfter|float|Gets or sets the amount of spacing, in grid lines. after the paragraph.|[1.1][]|
|lineUnitBefore|float|Gets or sets the amount of spacing, in grid lines, before the paragraph.|[1.1][]|
|outlineLevel|int|Gets or sets the outline level for the paragraph.|[1.1][]|
|rightIndent|float|Gets or sets the right indent value, in points, for the paragraph.|[1.1][]|
|spaceAfter|float|Gets or sets the spacing, in points, after the paragraph.|[1.1][]|
|spaceBefore|float|Gets or sets the spacing, in points, before the paragraph.|[1.1][]|
|style|string|Gets or sets the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.|[1.1][]|
|[styleBuiltIn](enums.md)|string|Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property. Possible values are: Other, Normal, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9, Toc1, more...|[1.3][]|
|tableNestingLevel|int|Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.|[1.3][]|
|text|string|Gets the text of the paragraph. Read-only.|[1.1][]|

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the paragraph. Read-only.|[1.1][]|
|font|[Font](font.md)|Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.|[1.1][]|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects in the paragraph. The collection does not include floating images. Read-only.|[1.1][]|
|list|[List](list.md)|Gets the List to which this paragraph belongs. Throws if the paragraph is not in a list. Read-only.|[1.3][]|
|listItem|[ListItem](listitem.md)|Gets the ListItem for the paragraph. Throws if the paragraph is not part of a list. Read-only.|[1.3][]|
|listItemOrNullObject|[ListItem](listitem.md)|Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.|[1.3][]|
|listOrNullObject|[List](list.md)|Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.|[1.3][]|
|parentBody|[Body](body.md)|Gets the parent body of the paragraph. Read-only.|[1.3][]|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the paragraph. Throws if there isn't a parent content control. Read-only.|[1.1][]|
|parentContentControlOrNullObject|[ContentControl](contentcontrol.md)|Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.|[1.3][]|
|parentTable|[Table](table.md)|Gets the table that contains the paragraph. Throws if it is not contained in a table. Read-only.|[1.3][]|
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the paragraph. Throws if it is not contained in a table cell. Read-only.|[1.3][]|
|parentTableCellOrNullObject|[TableCell](tablecell.md)|Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.|[1.3][]|
|parentTableOrNullObject|[Table](table.md)|Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.|[1.3][]|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[attachToList(listId: number, level: number)][]    |[List](list.md)|Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.|[1.3][]|
|[clear()][] |void|Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.|[1.1][]|
|[delete()][] |void|Deletes the paragraph and its content from the document.|[1.1][]|
|[detachFromList()][] |void|Moves this paragraph out of its list, if the paragraph is a list item.|[1.3][]|
|[getHtml()][] |string|Gets the HTML representation of the paragraph object.|[1.1][]|
|[getNext()][] |[Paragraph](paragraph.md)|Gets the next paragraph. Throws if the paragraph is the last one.|[1.3][]|
|[getNextOrNullObject()][] |[Paragraph](paragraph.md)|Gets the next paragraph. Returns a null object if the paragraph is the last one.|[1.3][]|
|[getOoxml()][] |string|Gets the Office Open XML (OOXML) representation of the paragraph object.|[1.1][]|
|[getPrevious()][] |[Paragraph](paragraph.md)|Gets the previous paragraph. Throws if the paragraph is the first one.|[1.3][]|
|[getPreviousOrNullObject()][] |[Paragraph](paragraph.md)|Gets the previous paragraph. Returns a null object if the paragraph is the first one.|[1.3][]|
|[getRange(rangeLocation: string)][] |[Range](range.md)|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|[1.3][]|
|[getTextRanges(endingMarks: string[], trimSpacing: bool)][getTextRanges] |[RangeCollection](rangecollection.md)|Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.|[1.3][]|
|[insertBreak(breakType: string, insertLocation: string)][] |void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.|[1.1][]|
|[insertContentControl()][] |[ContentControl](contentcontrol.md)|Wraps the paragraph object with a rich text content control.|[1.1][]|
|[insertFileFromBase64(base64File: string, insertLocation: string)][] |[Range](range.md)|Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|[1.1][]|
|[insertHtml(html: string, insertLocation: string)][] |[Range](range.md)|Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|[1.1][]|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)][] |[InlinePicture](inlinepicture.md)|Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|[1.1][]|
|[insertOoxml(ooxml: string, insertLocation: string)][] |[Range](range.md)|Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|[1.1][]|
|[insertParagraph(paragraphText: string, insertLocation: string)][] |[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|[1.1][]|
|[insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[])][insertTable] |[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|[1.3][]|
|[insertText(text: string, insertLocation: string)][] |[Range](range.md)|Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|[1.1][]|
|[load(param: object)][] |void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1][]|
|[search(searchText: string, searchOptions: SearchOptions)][search] |[RangeCollection](rangecollection.md)|Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.|[1.1][]|
|[select(selectionMode: string)][] |void|Selects and navigates the Word UI to the paragraph.|[1.1][]|
|[split(delimiters: string[], trimDelimiters: bool, trimSpacing: bool)][split] |[RangeCollection](rangecollection.md)|Splits the paragraph into child ranges by using delimiters.|[1.3][]|
|[startNewList()][] |[List](list.md)|Starts a new list with this paragraph. Fails if the paragraph is already a list item.|[1.3][]|


## Method Details


### attachToList(listId: number, level: number)
Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.

#### Syntax
```js
paragraphObject.attachToList(listId, level);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|listId|number|Required. The ID of an existing list.|
|level|number|Required. The level in the list.|

#### Returns
[List](list.md)

### clear()
Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.

#### Syntax
```js
paragraphObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to clear the contents of the first paragraph.
        paragraphs.items[0].clear();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Cleared the contents of the first paragraph.');
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


### delete()
Deletes the paragraph and its content from the document.

#### Syntax
```js
paragraphObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples

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

        // Queue a command to delete the first paragraph.
        paragraphs.items[0].delete();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Deleted the first paragraph.');
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


### detachFromList()
Moves this paragraph out of its list, if the paragraph is a list item.

#### Syntax
```js
paragraphObject.detachFromList();
```

#### Parameters
None

#### Returns
void

### getHtml()
Gets the HTML representation of the paragraph object.

#### Syntax
```js
paragraphObject.getHtml();
```

#### Parameters
None

#### Returns
string

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

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



### getNext()
Gets the next paragraph. Throws if the paragraph is the last one.

#### Syntax
```js
paragraphObject.getNext();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

### getNextOrNullObject()
Gets the next paragraph. Returns a null object if the paragraph is the last one.

#### Syntax
```js
paragraphObject.getNextOrNullObject();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

### getOoxml()
Gets the Office Open XML (OOXML) representation of the paragraph object.

#### Syntax
```js
paragraphObject.getOoxml();
```

#### Parameters
None

#### Returns
string

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the OOXML of the first paragraph.
        var ooxml = paragraphs.items[0].getOoxml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph OOXML: ' + ooxml.value);
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


### getPrevious()
Gets the previous paragraph. Throws if the paragraph is the first one.

#### Syntax
```js
paragraphObject.getPrevious();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)


### getPreviousOrNullObject()
Gets the previous paragraph. Returns a null object if the paragraph is the first one.

#### Syntax
```js
paragraphObject.getPreviousOrNullObject();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

#### Examples

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

        // Queue commands to create a proxy object for the next-to-last paragraph.
        var indexOfLastParagraph = paragraphs.items.length - 1;
        var precedingParagraph = paragraphs.items[indexOfLastParagraph].getPreviousOrNullObject();

        // Queue a command to load the text of the preceding paragraph.
        context.load(precedingParagraph, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            if (precedingParagraph.isNullObject) {
                console.log('There are no paragraphs before the current one.');
            } else {
                console.log('The preceding paragraph is: ' + precedingParagraph.text);
            }
        });
    });
}).catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```



### getRange(rangeLocation: string)
Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.

#### Syntax
```js
paragraphObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|string|Optional. Optional. The range location can be 'Whole', 'Start', 'End', 'After' or 'Content'.  Possible values are: Whole, Start, End, Before, After, Content|

#### Returns
[Range](range.md)

### getTextRanges(endingMarks: string[], trimSpacing: bool)
Gets the text ranges in the paragraph by using punctuation marks andor other ending marks.

#### Syntax
```js
paragraphObject.getTextRanges(endingMarks, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|endingMarks|string[]|Required. The punctuation marks and/or other ending marks as an array of strings.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)

### insertBreak(breakType: string, insertLocation: string)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|string|Required. The break type to add to the document. Possible values are: `Page` Page break at the insertion point.,`Column` Column break at the insertion point.,`Next` Section break on next page.,`SectionContinuous` New section without a corresponding page break.,`SectionEven` Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.,`SectionOdd` Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.,`Line` Line break.,`LineClearLeft` Line break.,`LineClearRight` Line break.,`TextWrapping` Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
void

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert a page break after the first paragraph.
        paragraph.insertBreak('page', 'After');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a page break after the paragraph.');
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


### insertContentControl()
Wraps the paragraph object with a rich text content control.

#### Syntax
```js
paragraphObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to wrap the first paragraph in a rich text content control.
        paragraph.insertContentControl();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Wrapped the first paragraph in a content control.');
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


### insertFileFromBase64(base64File: string, insertLocation: string)
Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: string)
Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted in the paragraph.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert HTML content at the end of the first paragraph.
        paragraph.insertHtml('<strong>Inserted HTML.</strong>', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted HTML content at the end of the first paragraph.');
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


### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)
Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[InlinePicture](inlinepicture.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        var b64encodedImg = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

        // Queue a command to insert a base64 encoded image at the beginning of the first paragraph.
        paragraph.insertInlinePictureFromBase64(b64encodedImg, Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added an image to the first paragraph.');
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


### insertOoxml(ooxml: string, insertLocation: string)
Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML to be inserted in the paragraph.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: string)
Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
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

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert the paragraph after the current paragraph.
        paragraph.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a new paragraph at the end of the first paragraph.');
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


### insertTable(rowCount: number, columnCount: number, insertLocation: string, values: string[][])
Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
paragraphObject.insertTable(rowCount, columnCount, insertLocation, values);
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
Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. Text to be inserted.|
|insertLocation|string|Required. The value can be 'Replace', 'Start' or 'End'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert text into the end of the paragraph.
        paragraph.insertText('New text inserted into the paragraph.', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted text at the end of the first paragraph.');
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

### search(searchText: string, searchOptions: SearchOptions)
Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.

#### Syntax
```js
paragraphObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|searchText|string|Required. The search text.|
|searchOptions|[SearchOptions]|Optional. Optional. Options for the search.|

#### Returns
[RangeCollection](rangecollection.md)

### select(selectionMode: string)
Selects and navigates the Word UI to the paragraph.

#### Syntax
```js
paragraphObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|string|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.  Possible values are: Select, Start, End|

#### Returns
void

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the last paragraph a create a
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1];

        // Queue a command to select the paragraph. The Word UI will
        // move to the selected paragraph.
        paragraph.select();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
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

### split(delimiters: string[], trimDelimiters: bool, trimSpacing: bool)
Splits the paragraph into child ranges by using delimiters.

#### Syntax
```js
paragraphObject.split(delimiters, trimDelimiters, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|delimiters|string[]|Required. The delimiters as an array of strings.|
|trimDelimiters|bool|Optional. Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)

### startNewList()
Starts a new list with this paragraph. Fails if the paragraph is already a list item.

#### Syntax
```js
paragraphObject.startNewList();
```

#### Parameters
None

#### Returns
[List](list.md)


[attachToList(listId: number, level: number)]: #attachtolistlistid-number-level-number
[clear()]: #clear
[delete()]: #delete
[detachFromList()]: #detachfromlist
[getHtml()]: #gethtml
[getNext()]: #getnext
[getNextOrNullObject()]: #getnextornullobject
[getOoxml()]: #getooxml
[getPrevious()]: #getprevious
[getPreviousOrNullObject()]: #getpreviousornullobject
[getRange(rangeLocation: string)]: #getrangerangelocation-string
[getTextRanges]: #gettextrangesendingmarks-string-trimspacing-bool
[insertBreak(breakType: string, insertLocation: string)]: #insertbreakbreaktype-string-insertlocation-string
[insertContentControl()]: #insertcontentcontrol
[insertFileFromBase64(base64File: string, insertLocation: string)]: #insertfilefrombase64base64file-string-insertlocation-string
[insertHtml(html: string, insertLocation: string)]: #inserthtmlhtml-string-insertlocation-string
[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)]: #insertinlinepicturefrombase64base64encodedimage-string-insertlocation-string
[insertOoxml(ooxml: string, insertLocation: string)]: #insertooxmlooxml-string-insertlocation-string
[insertParagraph(paragraphText: string, insertLocation: string)]: #insertparagraphparagraphtext-string-insertlocation-string
[insertTable]: #inserttablerowcount-number-columncount-number-insertlocation-string-values-string
[insertText(text: string, insertLocation: string)]: #inserttexttext-string-insertlocation-string
[load(param: object)]: #loadparam-object
[search]: #searchsearchtext-string-searchoptions-searchoptions
[select(selectionMode: string)]: #selectselectionmode-string
[split]: #splitdelimiters-string-trimdelimiters-bool-trimspacing-bool
[startNewList()]: #startnewlist
[SearchOptions]: searchoptions.md
[1.1]: ../requirement-sets/word-api-requirement-sets.md
[1.3]: ../requirement-sets/word-api-requirement-sets.md
