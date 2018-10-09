# Word JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Word add-ins run across multiple versions of Office, including Office 2016 or later for Windows, Office for iPad, Office for Mac, and Office Online. The following table lists the Word requirement sets, the Office host applications that support that requirement set, and the build or version numbers for those applications.

> [!NOTE]
> For the requirement sets that are marked as Beta, use the specified (or later) version of the Office software and use the Beta library of the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> 
> Entries not listed as Beta are generally available and you can continue to use Production CDN library: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Requirement set  |   Office 365 for Windows\*  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| WordApi 1.3 | Version 1612 (Build 7668.1000) or later| March 2017, 2.22 or later | March 2017, 15.32 or later| March 2017 ||
| WordApi 1.2  | December 2015 update, Version 1601 (Build 6568.1000) or later | January 2016, 1.18 or later | January 2016, 15.19 or later| September 2016 | |
| WordApi 1.1  | Version 1509 (Build 4266.1001) or later| January 2016, 1.18 or later | January 2016, 15.19 or later| September 2016 | |

> [!NOTE]
> The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the WordApi 1.1 requirement set.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## Office common API requirement sets

For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## What's new in Word JavaScript API 1.3 

The following are the new additions to the Word JavaScript APIs in requirement set 1.3. 

|Object| What's new| Description|Requirement set| 
|:-----|-----|:----|:----| 
|[application](/javascript/api/word/word.application)|_Method_ > createDocument(base64File: string) | Creates a new document by using a base64 encoded .docx file. Read-only.|1.3|
|[body](/javascript/api/word/word.body)|_Relationship_ > lists|Gets the collection of list objects in the body. Read-only.|1.3|
|[body](/javascript/api/word/word.body)|_Relationship_ > parentBody|Gets the parent body of the body. For example, a table cell body's parent body could be a header. Read-only.|1.3|
|[body](/javascript/api/word/word.body)|_Relationship_ > parentSection|Gets the parent section of the body. Read-only.|1.3|
|[body](/javascript/api/word/word.body)|_Relationship_ > styleBuiltIn|Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|1.3|
|[body](/javascript/api/word/word.body)|_Relationship_ > tables|Gets the collection of table objects in the body. Read-only.|1.3|
|[body](/javascript/api/word/word.body)|_Relationship_ > type|Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.|1.3|
|[body](/javascript/api/word/word.body)|_Method_ > getRange(rangeLocation: RangeLocation)|Gets the whole body, or the starting or ending point of the body, as a range.|1.3|
|[body](/javascript/api/word/word.body)|_Method_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_Relationship_ > breaks|Specifies the form of a break: line, page, or section type. Read-only.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relationship_ > lists|Gets the collection of list objects in the content control. Read-only.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relationship_ > parentBody|Gets the parent body of the content control. Read-only.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relationship_ > parentTable|Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relationship_ > parentTableCell|Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relationship_ > styleBuiltIn|Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relationship_ > subtype|Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relationship_ > tables|Gets the collection of table objects in the content control. Read-only.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Method_ > getRange(rangeLocation: RangeLocation)|Gets the whole content control, or the starting or ending point of the content control, as a range.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Method_ > getTextRanges(endingMarks: string, trimSpacing: bool)|Gets the text ranges in the content control by using punctuation marks andor other ending marks.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Method_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Method_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Splits the content control into child ranges by using delimiters.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Method_ > getByTypes(types: ContentControlType)|Gets the content controls that have the specified types andor subtypes.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Method_ > getFirst()|Gets the first content control in this collection.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Property_ > key|Gets the key of the custom property. Read only. |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Property_ > value|Gets or sets the value of the custom property.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Relationship_ > type|Gets the value type of the custom property. Read-only.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Method_ > delete()|Deletes the custom property.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Property_ > items|A collection of customProperty objects. Read-only.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Method_ > deleteAll()|Deletes all custom properties in this collection.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Method_ > getCount()|Gets the count of custom properties.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Method_ > getItem(key: string)|Gets a custom property object by its key, which is case-insensitive.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Method_ > set(key: string, value: object)|Creates or sets a custom property.|1.3|
|[document](/javascript/api/word/word.document)|_Relationship_ > properties|Gets the properties of the current document. Read-only.|1.3|
|[document](/javascript/api/word/word.document)|_Method_ > open()|Open the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > applicationName|Gets the application name of the document. Read only.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > author|Gets or sets the author of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > category|Gets or sets the category of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > comments|Gets or sets the comments of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > company|Gets or sets the company of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > format|Gets or sets the format of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > keywords|Gets or sets the keywords of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > lastAuthor|Gets or sets the last author of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > manager|Gets or sets the manager of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > revisionNumber|Gets the revision number of the document. Read-only.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > security|Gets the security of the document. Read-only.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > subject|Gets or sets the subject of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > template|Gets the template of the document. Read-only.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Property_ > title|Gets or sets the title of the document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relationship_ > creationDate|Gets the creation date of the document. Read-only.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relationship_ > customProperties|Gets the collection of custom properties of the document Read-only.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relationship_ > lastPrintDate|Gets the last print date of the document. Read-only.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relationship_ > lastSaveTime|Gets the last save time of the document. Read only.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relationship_ > parentTable|Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relationship_ > parentTableCell|Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > getNext()|Gets the next inline image.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > getRange(rangeLocation: RangeLocation)|Gets the picture, or the starting or ending point of the picture, as a range.|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_Method_ > getFirst()|Gets the first inline image in this collection.|1.3|
|[list](/javascript/api/word/word.list)|_Property_ > id|Gets the list's id. Read-only.|1.3|
|[list](/javascript/api/word/word.list)|_Property_ > levelExistences|Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.|1.3|
|[list](/javascript/api/word/word.list)|_Relationship_ > levelTypes|Gets all 9 level types in the list. Each type can be 'Bullet', 'Number' or 'Picture'. Read-only.|1.3|
|[list](/javascript/api/word/word.list)|_Relationship_ > paragraphs|Gets paragraphs in the list. Read-only.|1.3|
|[list](/javascript/api/word/word.list)|_Method_ > getLevelParagraphs(level: number)|Gets the paragraphs that occur at the specified level in the list.|1.3|
|[list](/javascript/api/word/word.list)|_Method_ > getLevelString(level: number)|Gets the bullet, number or picture at the specified level as a string.|1.3|
|[list](/javascript/api/word/word.list)|_Method_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|1.3|
|[list](/javascript/api/word/word.list)|_Method_ > setLevelAlignment(level: number, alignment: Alignment)|Sets the alignment of the bullet, number or picture at the specified level in the list.|1.3|
|[list](/javascript/api/word/word.list)|_Method_ > setLevelBullet(level: number, listBullet: ListBullet, charCode: number, fontName: string)|Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.|1.3|
|[list](/javascript/api/word/word.list)|_Method_ > setLevelIndents(level: number, textIndent: float, textIndent: float)|Sets the two indents of the specified level in the list.|1.3|
|[list](/javascript/api/word/word.list)|_Method_ > setLevelNumbering(level: number, listNumbering: ListNumbering, formatString: object)|Sets the numbering format at the specified level in the list.|1.3|
|[list](/javascript/api/word/word.list)|_Method_ > setLevelStartingNumber(level: number, startingNumber: number)|Sets the starting number at the specified level in the list. Default value is 1.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Property_ > items|A collection of list objects. Read-only.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Method_ > getById(id: number)|Gets a list by its identifier.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Method_ > getFirst()|Gets the first list in this collection.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Method_ > getItem(index: number)|Gets a list object by its index in the collection.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Property_ > level|Gets or sets the level of the item in the list.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Property_ > listString|Gets the list item bullet, number or picture as a string. Read-only.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Property_ > siblingIndex|Gets the list item order number in relation to its siblings. Read-only.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Method_ > getAncestor(parentOnly: bool)|Gets the list item parent, or the closest ancestor if the parent does not exist.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Method_ > getDescendants(directChildrenOnly: bool)|Gets all descendant list items of the list item.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Property_ > isLastParagraph|Indicates the paragraph is the last one inside its parent body. Read-only.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Property_ > isListItem|Checks whether the paragraph is a list item. Read-only.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Property_ > tableNestingLevel|Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relationship_ > list|Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relationship_ > listItem|Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relationship_ > parentBody|Gets the parent body of the paragraph. Read-only.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relationship_ > parentTable|Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relationship_ > parentTableCell|Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relationship_ > styleBuiltIn|Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > attachToList(listId: number, level: number)|Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > detachFromList()|Moves this paragraph out of its list, if the paragraph is a list item.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > getNext()|Gets the next paragraph.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > getPrevious()|Gets the previous paragraph.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > getRange(rangeLocation: RangeLocation)|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > getTextRanges(endingMarks: string, trimSpacing: bool)|Gets the text ranges in the paragraph by using punctuation marks andor other ending marks.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > split(delimiters: string, trimDelimiters: bool, trimSpacing: bool)|Splits the paragraph into child ranges by using delimiters.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Method_ > startNewList()|Starts a new list with this paragraph. Fails if the paragraph is already a list item.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Method_ > getFirst()|Gets the first paragraph in this collection.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Method_ > getLast()|Gets the last paragraph in this collection.|1.3|
|[range](/javascript/api/word/word.range)|_Property_ > hyperlink|Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a newline character ('\n') to separate the address part from the optional location part.|1.3|
|[range](/javascript/api/word/word.range)|_Property_ > isEmpty|Checks whether the range length is zero. Read-only.|1.3|
|[range](/javascript/api/word/word.range)|_Relationship_ > lists|Gets the collection of list objects in the range. Read-only.|1.3|
|[range](/javascript/api/word/word.range)|_Relationship_ > parentBody|Gets the parent body of the range. Read-only.|1.3|
|[range](/javascript/api/word/word.range)|_Relationship_ > parentTable|Gets the table that contains the range. Returns null if it is not contained in a table. Read-only.|1.3|
|[range](/javascript/api/word/word.range)|_Relationship_ > parentTableCell|Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.|1.3|
|[range](/javascript/api/word/word.range)|_Relationship_ > styleBuiltIn|Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|1.3|
|[range](/javascript/api/word/word.range)|_Relationship_ > tables|Gets the collection of table objects in the range. Read-only.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > compareLocationWith(range: Range)|Compares this range's location with another range's location.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > expandTo(range: Range)|Returns a new range that extends from this range in either direction to cover another range. This range is not changed.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > getHyperlinkRanges()|Gets hyperlink child ranges within the range.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > getNextTextRange(endingMarks: string, trimSpacing: bool)|Gets the next text range by using punctuation marks andor other ending marks.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > getRange(rangeLocation: RangeLocation)|Clones the range, or gets the starting or ending point of the range as a new range.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > getTextRanges(endingMarks: string, trimSpacing: bool)|Gets the text child ranges in the range by using punctuation marks andor other ending marks.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > intersectWith(range: Range)|Returns a new range as the intersection of this range with another range. This range is not changed.|1.3|
|[range](/javascript/api/word/word.range)|_Method_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Splits the range into child ranges by using delimiters.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Property_ > items|A collection of range objects. Read-only.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Method_ > getFirst()|Gets the first range in this collection.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Method_ > getItem(index: number)|Gets a range object by its index in the collection.|1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Method_ > load(object: object, option: object)|Fills the proxy object created in JavaScript layer with property and options specified in the parameter. |1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Method_ > sync()|Submits the request queue to Word and returns a promise object, which can be used for chaining further actions.|1.3|
|[section](/javascript/api/word/word.section)|_Method_ > getNext()|Gets the next section.|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_Method_ > getFirst()|Gets the first section in this collection.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > headerRowCount|Gets and sets the number of header rows.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > height|Gets the height of the table in points. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > isUniform|Indicates whether all of the table rows are uniform. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > nestingLevel|Gets the nesting level of the table. Top-level tables have level 1. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > rowCount|Gets the number of rows in the table. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > shadingColor|Gets and sets the shading color.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > style|Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > styleBandedColumns|Gets and sets whether the table has banded columns.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > styleBandedRows|Gets and sets whether the table has banded rows.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > styleFirstColumn|Gets and sets whether the table has a first column with a special style.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > styleLastColumn|Gets and sets whether the table has a last column with a special style.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > styleTotalRow|Gets and sets whether the table has a total (last) row with a special style.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > values|Gets and sets the text values in the table, as a 2D Javascript array.|1.3|
|[table](/javascript/api/word/word.table)|_Property_ > width|Gets and sets the width of the table in points.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > font|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > horizontalAlignment|Gets and sets the horizontal alignment of every cell in the table. The value can be 'left', 'centered', 'right', or 'justified'.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > paragraphAfter|Gets the paragraph after the table. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > paragraphBefore|Gets the paragraph before the table. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > parentBody|Gets the parent body of the table. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > parentContentControl|Gets the content control that contains the table. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > parentTable|Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > parentTableCell|Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > rows|Gets all of the table rows. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > styleBuiltIn|Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > tables|Gets the child tables nested one level deeper. Read-only.|1.3|
|[table](/javascript/api/word/word.table)|_Relationship_ > verticalAlignment|Gets and sets the vertical alignment of every cell in the table. The value can be 'top', 'center' or 'bottom'.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > addColumns(insertLocation: InsertLocation, columnCount: number, values: string)|Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > addRows(insertLocation: InsertLocation, rowCount: number, values: string)|Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > autoFitContents()|Autofits the table columns to the width of their contents.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > autoFitWindow()|Autofits the table columns to the width of the window.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > clear()|Clears the contents of the table.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > delete()|Deletes the entire table.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > deleteColumns(columnIndex: number, columnCount: number)|Deletes specific columns. This is applicable to uniform tables.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > deleteRows(rowIndex: number, rowCount: number)|Deletes specific rows.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > distributeColumns()|Distributes the column widths evenly.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > distributeRows()|Distributes the row heights evenly.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > getBorder(borderLocation: BorderLocation)|Gets the border style for the specified border.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > getCell(rowIndex: number, cellIndex: number)|Gets the table cell at a specified row and column.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Gets cell padding in points.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > getNext()|Gets the next table.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > getRange(rangeLocation: RangeLocation)|Gets the range that contains this table, or the range at the start or end of the table.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > insertContentControl()|Inserts a content control on the table.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > select(selectionMode: SelectionMode)|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|1.3|
|[table](/javascript/api/word/word.table)|_Method_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Sets cell padding in points.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Property_ > color|Gets or sets the table border color, as a hex value or name.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Property_ > width|Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Relationship_ > type|Gets or sets the type of the table border.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Property_ > cellIndex|Gets the index of the cell in its row. Read-only.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Property_ > columnWidth|Gets and sets the width of the cell's column in points. This is applicable to uniform tables.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Property_ > rowIndex|Gets the index of the cell's row in the table. Read-only.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Property_ > shadingColor|Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Property_ > value|Gets and sets the text of the cell.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Property_ > width|Gets the width of the cell in points. Read-only.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relationship_ > body|Gets the body object of the cell. Read-only.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relationship_ > horizontalAlignment|Gets and sets the horizontal alignment of the cell. The value can be 'left', 'centered', 'right', or 'justified'.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relationship_ > parentRow|Gets the parent row of the cell. Read-only.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relationship_ > parentTable|Gets the parent table of the cell. Read-only.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relationship_ > verticalAlignment|Gets and sets the vertical alignment of the cell. The value can be 'top', 'center' or 'bottom'.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Method_ > deleteColumn()|Deletes the column containing this cell. This is applicable to uniform tables.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Method_ > deleteRow()|Deletes the row containing this cell.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Method_ > getBorder(borderLocation: BorderLocation)|Gets the border style for the specified border.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Method_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Gets cell padding in points.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Method_ > getNext()|Gets the next cell.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Method_ > insertColumns(insertLocation: InsertLocation, columnCount: number, values: string)|Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Method_ > insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Method_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Sets cell padding in points.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Property_ > items|A collection of tableCell objects. Read-only.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Method_ > getFirst()|Gets the first table cell in this collection.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Method_ > getItem(index: number)|Gets a table cell object by its index in the collection.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Property_ > items|A collection of table objects. Read-only.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Method_ > getFirst()|Gets the first table in this collection.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Method_ > getItem(index: number)|Gets a table object by its index in the collection.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Property_ > cellCount|Gets the number of cells in the row. Read-only.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Property_ > isHeader|Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object. Read-only.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Property_ > preferredHeight|Gets and sets the preferred height of the row in points.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Property_ > rowIndex|Gets the index of the row in its parent table. Read-only.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Property_ > shadingColor|Gets and sets the shading color.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Property_ > values|Gets and sets the text values in the row, as a 1D Javascript array.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relationship_ > cells|Gets cells. Read-only.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relationship_ > font|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relationship_ > horizontalAlignment|Gets and sets the horizontal alignment of every cell in the row. The value can be 'left', 'centered', 'right', or 'justified'.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relationship_ > parentTable|Gets parent table. Read-only.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relationship_ > verticalAlignment|Gets and sets the vertical alignment of the cells in the row. The value can be 'top', 'center' or 'bottom'.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > clear()|Clears the contents of the row.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > delete()|Deletes the entire row.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > getBorder(borderLocation: BorderLocation)|Gets the border style of the cells in the row.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Gets cell padding in points.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > getNext()|Gets the next row.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > select(selectionMode: SelectionMode)|Selects the row and navigates the Word UI to it.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Method_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Sets cell padding in points.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Property_ > items|A collection of tableRow objects. Read-only.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Method_ > getFirst()|Gets the first row in this collection.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Method_ > getItem(index: number)|Gets a table row object by its index in the collection.|1.3|


## What's new in Word JavaScript API 1.2

The following are the new additions to the Word JavaScript APIs in requirement set 1.2. 

|Object| What's new| Description|Requirement set|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Method_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relationship_ > paragraph|Gets the parent paragraph that contains the inline image. Read-only.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > delete()|Deletes the inline picture from the document.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > insertBreak(breakType: BreakType, insertLocation: InsertLocation)|Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > insertFileFromBase64(base64File: string, insertLocation: InsertLocation)|Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > insertHtml(html: string, insertLocation: InsertLocation)|Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation|Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > insertOoxml(ooxml: string, insertLocation: InsertLocation)|Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > insertText(text: string, insertLocation: InsertLocation)|Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Method_ > select(selectionMode: SelectionMode)|Selects the inline picture. This causes Word to scroll to the selection.|1.2|
|[range](/javascript/api/word/word.range)|_Relationship_ > inlinePictures|Gets the collection of inline picture objects in the range. Read-only.|1.2|
|[range](/javascript/api/word/word.range)|_Method_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|Inserts a picture at the specified location. The insertLocation value can be 'Replace', 'Start', 'End', 'Before' or 'After'.|1.2|

## Word JavaScript API 1.1

Word JavaScript API 1.1 is the first version of the API. For details about the API,  see the [Word JavaScript API](/javascript/api/word) reference topics. 

## See also

- [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
