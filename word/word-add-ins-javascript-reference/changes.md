|Resource|Member Type|Name|Data type |Description|Requirement-Set|
|:----|:----|:----|:----|:----|:----|
|[application](application.md)|Method|[createDocument(base64File: string)](application.md#createdocumentbase64file-string)|[Document](document.md)|Creates a new document by using a base64 encoded .docx file.|WordApiDesktop, 1.3|
|[body](body.md)|Relationship|lists|[ListCollection](listcollection.md)|Gets the collection of list objects in the body. Read-only.|1.3|
|[body](body.md)|Relationship|parentBody|[Body](body.md)|Gets the parent body of the body. For example, a table cell body's parent body could be a header. Read-only.|1.3|
|[body](body.md)|Relationship|tables|[TableCollection](tablecollection.md)|Gets the collection of table objects in the body. Read-only.|1.3|
|[body](body.md)|Relationship|type|[BodyType](bodytype.md)|Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.|1.3|
|[body](body.md)|Method|[getRange(rangeLocation: RangeLocation)](body.md#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the whole body, or the starting or ending point of the body, as a range.|1.3|
|[body](body.md)|Method|[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](body.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.|1.3|
|[contentControl](contentcontrol.md)|Relationship|lists|[ListCollection](listcollection.md)|Gets the collection of list objects in the content control. Read-only.|1.3|
|[contentControl](contentcontrol.md)|Relationship|parentTable|[Table](table.md)|Gets the table that contains the content control. Returns null if it is not contained in a table. Read-only.|1.3|
|[contentControl](contentcontrol.md)|Relationship|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the content control. Returns null if it is not contained in a table cell. Read-only.|1.3|
|[contentControl](contentcontrol.md)|Relationship|subtype|[ContentControlType](contentcontroltype.md)|Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.|1.3|
|[contentControl](contentcontrol.md)|Relationship|tables|[TableCollection](tablecollection.md)|Gets the collection of table objects in the content control. Read-only.|1.3|
|[contentControl](contentcontrol.md)|Method|[getRange(rangeLocation: RangeLocation)](contentcontrol.md#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the whole content control, or the starting or ending point of the content control, as a range.|1.3|
|[contentControl](contentcontrol.md)|Method|[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](contentcontrol.md#gettextrangespunctuationmarks-string-trimspacing-bool)|[RangeCollection](rangecollection.md)|Gets the text ranges in the content control by using punctuation marks andor space character.|1.3|
|[contentControl](contentcontrol.md)|Method|[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](contentcontrol.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|1.3|
|[contentControl](contentcontrol.md)|Method|[split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)](contentcontrol.md#splitdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimspacing-bool)|[RangeCollection](rangecollection.md)|Splits the content control into child ranges by using delimiters.|1.3|
|[contentControlCollection](contentcontrolcollection.md)|Method|[getByTypes(types: ContentControlType[])](contentcontrolcollection.md#getbytypestypes-contentcontroltype)|[ContentControlCollection](contentcontrolcollection.md)|Gets the content controls that have the specified types andor subtypes.|WordApiDesktop, 1.3|
|[document](document.md)|Method|[open()](document.md#open)|void|Open the document.|WordApiDesktop, 1.3|
|[font](font.md)|Property|doubleStrikeThrough|bool|Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.|WordApiDesktop, 1.3|
|[inlinePicture](inlinepicture.md)|Relationship|imageFormat|[ImageFormat](imageformat.md)|Gets the format of the inline image. Read-only.|1.3|
|[inlinePicture](inlinepicture.md)|Relationship|next|[InlinePicture](inlinepicture.md)|Gets the next inline image. Read-only.|1.3|
|[inlinePicture](inlinepicture.md)|Relationship|parentTable|[Table](table.md)|Gets the table that contains the inline image. Returns null if it is not contained in a table. Read-only.|1.3|
|[inlinePicture](inlinepicture.md)|Relationship|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the inline image. Returns null if it is not contained in a table cell. Read-only.|1.3|
|[inlinePicture](inlinepicture.md)|Method|[getRange(rangeLocation: RangeLocation)](inlinepicture.md#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the picture, or the starting or ending point of the picture, as a range.|1.3|
|[inlinePictureCollection](inlinepicturecollection.md)|Relationship|first|[InlinePicture](inlinepicture.md)|Gets the first inline image in this collection. Read-only.|1.3|
|[list](list.md)|Property|id|int|Gets the list's id. Read-only.|1.3|
|[list](list.md)|Relationship|paragraphs|[ParagraphCollection](paragraphcollection.md)|A collection containing the paragraphs in this list. Read-only.|1.3|
|[list](list.md)|Method|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](list.md#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|1.3|
|[listCollection](listcollection.md)|Property|items|[List[]](list.md)|A collection of list objects. Read-only.|1.3|
|[listCollection](listcollection.md)|Relationship|first|[List](list.md)|Gets the first list in this collection. Read-only.|1.3|
|[listCollection](listcollection.md)|Method|[getById(id: number)](listcollection.md#getbyidid-number)|[List](list.md)|Gets a list by its identifier.|1.3|
|[listCollection](listcollection.md)|Method|[getItem(index: number)](listcollection.md#getitemindex-number)|[List](list.md)|Gets a list object by its index in the collection.|1.3|
|[paragraph](paragraph.md)|Property|listLevel|int|Gets or sets the list level of the paragraph.|1.3|
|[paragraph](paragraph.md)|Property|outlineLevel|int|Gets or sets the outline level for the paragraph.|WordApiDesktop, 1.3|
|[paragraph](paragraph.md)|Property|tableNestingLevel|int|Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.|1.3|
|[paragraph](paragraph.md)|Relationship|list|[List](list.md)|Gets the List to which this paragraph belongs. Returns null if the paragraph is not in a list. Read-only.|1.3|
|[paragraph](paragraph.md)|Relationship|next|[Paragraph](paragraph.md)|Gets the next paragraph. Read-only.|1.3|
|[paragraph](paragraph.md)|Relationship|parentBody|[Body](body.md)|Gets the parent body of the paragraph. Read-only.|1.3|
|[paragraph](paragraph.md)|Relationship|parentTable|[Table](table.md)|Gets the table that contains the paragraph. Returns null if it is not contained in a table. Read-only.|1.3|
|[paragraph](paragraph.md)|Relationship|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the paragraph. Returns null if it is not contained in a table cell. Read-only.|1.3|
|[paragraph](paragraph.md)|Relationship|previous|[Paragraph](paragraph.md)|Gets the previous paragraph. Read-only.|1.3|
|[paragraph](paragraph.md)|Method|[getRange(rangeLocation: RangeLocation)](paragraph.md#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|1.3|
|[paragraph](paragraph.md)|Method|[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](paragraph.md#gettextrangespunctuationmarks-string-trimspacing-bool)|[RangeCollection](rangecollection.md)|Gets the text ranges in the paragraph by using punctuation marks andor space character.|1.3|
|[paragraph](paragraph.md)|Method|[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](paragraph.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|1.3|
|[paragraph](paragraph.md)|Method|[split(delimiters: string[], trimDelimiters: bool, trimSpacing: bool)](paragraph.md#splitdelimiters-string-trimdelimiters-bool-trimspacing-bool)|[RangeCollection](rangecollection.md)|Splits the paragraph into child ranges by using delimiters.|1.3|
|[paragraphCollection](paragraphcollection.md)|Relationship|first|[Paragraph](paragraph.md)|Gets the first paragraph in this collection. Read-only.|1.3|
|[paragraphCollection](paragraphcollection.md)|Relationship|last|[Paragraph](paragraph.md)|Gets the last paragraph in this collection. Read-only.|1.3|
|[range](range.md)|Property|hyperlink|string|Gets the first hyperlink in the range, or sets a hyperlink on the range. Existing hyperlinks in this range are deleted when you set a new hyperlink.|1.3|
|[range](range.md)|Property|isEmpty|bool|Checks whether the range length is zero. Read-only.|1.3|
|[range](range.md)|Relationship|lists|[ListCollection](listcollection.md)|Gets the collection of list objects in the range. Read-only.|1.3|
|[range](range.md)|Relationship|parentBody|[Body](body.md)|Gets the parent body of the range. Read-only.|1.3|
|[range](range.md)|Relationship|parentTable|[Table](table.md)|Gets the table that contains the range. Returns null if it is not contained in a table. Read-only.|1.3|
|[range](range.md)|Relationship|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the range. Returns null if it is not contained in a table cell. Read-only.|1.3|
|[range](range.md)|Relationship|tables|[TableCollection](tablecollection.md)|Gets the collection of table objects in the range. Read-only.|1.3|
|[range](range.md)|Method|[compareLocationWith(range: Range)](range.md#comparelocationwithrange-range)|[LocationRelation](locationrelation.md)|Compares this range's location with another range's location.|1.3|
|[range](range.md)|Method|[expandTo(range: Range)](range.md#expandtorange-range)|void|Expands the range in either direction to cover another range.|1.3|
|[range](range.md)|Method|[getHyperlinkRanges()](range.md#gethyperlinkranges)|[RangeCollection](rangecollection.md)|Gets hyperlink child ranges within the range.|1.3|
|[range](range.md)|Method|[getNextTextRange(punctuationMarks: string[], trimSpacing: bool)](range.md#getnexttextrangepunctuationmarks-string-trimspacing-bool)|[Range](range.md)|Gets the next text range by using punctuation marks andor space character.|1.3|
|[range](range.md)|Method|[getRange(rangeLocation: RangeLocation)](range.md#getrangerangelocation-rangelocation)|[Range](range.md)|Clones the range, or gets the starting or ending point of the range as a new range.|1.3|
|[range](range.md)|Method|[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](range.md#gettextrangespunctuationmarks-string-trimspacing-bool)|[RangeCollection](rangecollection.md)|Gets the text child ranges in the range by using punctuation marks andor space character.|1.3|
|[range](range.md)|Method|[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](range.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|1.3|
|[range](range.md)|Method|[intersectWith(range: Range)](range.md#intersectwithrange-range)|void|Shrinks the range to the intersection of the range with another range.|1.3|
|[range](range.md)|Method|[split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)](range.md#splitdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimspacing-bool)|[RangeCollection](rangecollection.md)|Splits the range into child ranges by using delimiters.|1.3|
|[rangeCollection](rangecollection.md)|Property|items|[Range[]](range.md)|A collection of range objects. Read-only.|1.3|
|[rangeCollection](rangecollection.md)|Relationship|first|[Range](range.md)|Gets the first range in this collection. Read-only.|1.3|
|[rangeCollection](rangecollection.md)|Method|[getItem(index: number)](rangecollection.md#getitemindex-number)|[Range](range.md)|Gets a range object by its index in the collection.|1.3|
|[searchResultCollection](searchresultcollection.md)|Relationship|first|[Range](range.md)|Gets the first searched result in this collection. Read-only.|1.3|
|[section](section.md)|Relationship|next|[Section](section.md)|Gets the next section. Read-only.|1.3|
|[sectionCollection](sectioncollection.md)|Relationship|first|[Section](section.md)|Gets the first section in this collection. Read-only.|1.3|
|[table](table.md)|Property|headerRowCount|int|Gets and sets the number of header rows.|1.3|
|[table](table.md)|Property|isUniform|bool|Indicates whether all of the table rows are uniform. Read-only.|1.3|
|[table](table.md)|Property|nestingLevel|int|Gets the nesting level of the table. Top-level tables have level 1. Read-only.|1.3|
|[table](table.md)|Property|rowCount|int|Gets the number of rows in the table. Read-only.|1.3|
|[table](table.md)|Property|shadingColor|string|Gets and sets the shading color.|1.3|
|[table](table.md)|Property|style|string|Gets and sets the name of the table style.|1.3|
|[table](table.md)|Property|styleBandedColumns|bool|Gets and sets whether the table has banded columns.|1.3|
|[table](table.md)|Property|styleBandedRows|bool|Gets and sets whether the table has banded rows.|1.3|
|[table](table.md)|Property|styleFirstColumn|bool|Gets and sets whether the table has a first column with a special style.|1.3|
|[table](table.md)|Property|styleLastColumn|bool|Gets and sets whether the table has a last column with a special style.|1.3|
|[table](table.md)|Property|styleTotalRow|bool|Gets and sets whether the table has a total (last) row with a special style.|1.3|
|[table](table.md)|Property|values|string|Gets and sets the text values in the table, as a 2D Javascript array.|1.3|
|[table](table.md)|Relationship|cellPaddingBottom|[float](float.md)|Gets and sets the default bottom cell padding in points.|1.3|
|[table](table.md)|Relationship|cellPaddingLeft|[float](float.md)|Gets and sets the default left cell padding in points.|1.3|
|[table](table.md)|Relationship|cellPaddingRight|[float](float.md)|Gets and sets the default right cell padding in points.|1.3|
|[table](table.md)|Relationship|cellPaddingTop|[float](float.md)|Gets and sets the default top cell padding in points.|1.3|
|[table](table.md)|Relationship|font|[Font](font.md)|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|1.3|
|[table](table.md)|Relationship|height|[float](float.md)|Gets the height of the table in points. Read-only.|1.3|
|[table](table.md)|Relationship|next|[Table](table.md)|Gets the next table. Read-only.|1.3|
|[table](table.md)|Relationship|paragraphAfter|[Paragraph](paragraph.md)|Gets the paragraph after the table. Read-only.|1.3|
|[table](table.md)|Relationship|paragraphBefore|[Paragraph](paragraph.md)|Gets the paragraph before the table. Read-only.|1.3|
|[table](table.md)|Relationship|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the table. Read-only.|1.3|
|[table](table.md)|Relationship|parentTable|[Table](table.md)|Gets the table that contains this table. Returns null if it is not contained in a table. Read-only.|1.3|
|[table](table.md)|Relationship|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains this table. Returns null if it is not contained in a table cell. Read-only.|1.3|
|[table](table.md)|Relationship|rows|[TableRowCollection](tablerowcollection.md)|Gets all of the table rows. Read-only.|1.3|
|[table](table.md)|Relationship|tables|[TableCollection](tablecollection.md)|Gets the child tables nested one level deeper. Read-only.|1.3|
|[table](table.md)|Relationship|verticalAlignment|[VerticalAlignment](verticalalignment.md)|Gets and sets the vertical alignment of every cell in the table.|1.3|
|[table](table.md)|Relationship|width|[float](float.md)|Gets and sets the width of the table in points.|1.3|
|[table](table.md)|Method|[addColumns(insertLocation: InsertLocation, columnCount: number, values: string[][])](table.md#addcolumnsinsertlocation-insertlocation-columncount-number-values-string)|void|Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|1.3|
|[table](table.md)|Method|[addRows(insertLocation: InsertLocation, rowCount: number, values: string[][])](table.md#addrowsinsertlocation-insertlocation-rowcount-number-values-string)|void|Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.|1.3|
|[table](table.md)|Method|[autoFitContents()](table.md#autofitcontents)|void|Autofits the table columns to the width of their contents.|1.3|
|[table](table.md)|Method|[autoFitWindow()](table.md#autofitwindow)|void|Autofits the table columns to the width of the window.|1.3|
|[table](table.md)|Method|[clear()](table.md#clear)|void|Clears the contents of the table.|1.3|
|[table](table.md)|Method|[delete()](table.md#delete)|void|Deletes the entire table.|1.3|
|[table](table.md)|Method|[deleteColumns(columnIndex: number, columnCount: number)](table.md#deletecolumnscolumnindex-number-columncount-number)|void|Deletes specific columns. This is applicable to uniform tables.|1.3|
|[table](table.md)|Method|[deleteRows(rowIndex: number, rowCount: number)](table.md#deleterowsrowindex-number-rowcount-number)|void|Deletes specific rows.|1.3|
|[table](table.md)|Method|[distributeColumns()](table.md#distributecolumns)|void|Distributes the column widths evenly.|1.3|
|[table](table.md)|Method|[distributeRows()](table.md#distributerows)|void|Distributes the row heights evenly.|1.3|
|[table](table.md)|Method|[getBorderStyle(borderLocation: BorderLocation)](table.md#getborderstyleborderlocation-borderlocation)|[TableBorderStyle](tableborderstyle.md)|Gets the border style for the specified border.|1.3|
|[table](table.md)|Method|[getCell(rowIndex: number, cellIndex: number)](table.md#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Gets the table cell at a specified row and column.|1.3|
|[table](table.md)|Method|[getRange(rangeLocation: RangeLocation)](table.md#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the range that contains this table, or the range at the start or end of the table.|1.3|
|[table](table.md)|Method|[insertContentControl()](table.md#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Inserts a content control on the table.|1.3|
|[table](table.md)|Method|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](table.md#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|1.3|
|[table](table.md)|Method|[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](table.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|1.3|
|[table](table.md)|Method|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](table.md#mergecellstoprow-number-firstcell-number-bottomrow-number-lastcell-number)|[TableCell](tablecell.md)|Merges the cells bounded inclusively by a first and last cell.|WordApiDesktop, 1.3|
|[table](table.md)|Method|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](table.md#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.|1.3|
|[table](table.md)|Method|[select(selectionMode: SelectionMode)](table.md#selectselectionmode-selectionmode)|void|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|1.3|
|[tableBorderStyle](tableborderstyle.md)|Property|color|string|Gets or sets the table border color, as a hex value or name.|1.3|
|[tableBorderStyle](tableborderstyle.md)|Relationship|type|[BorderType](bordertype.md)|Gets or sets the type of the table border style.|1.3|
|[tableBorderStyle](tableborderstyle.md)|Relationship|width|[float](float.md)|Gets or sets the width, in points, of the table border style.|1.3|
|[tableCell](tablecell.md)|Property|cellIndex|int|Gets the index of the cell in its row. Read-only.|1.3|
|[tableCell](tablecell.md)|Property|rowIndex|int|Gets the index of the cell's row in the table. Read-only.|1.3|
|[tableCell](tablecell.md)|Property|shadingColor|string|Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.|1.3|
|[tableCell](tablecell.md)|Property|value|string|Gets and sets the text of the cell.|1.3|
|[tableCell](tablecell.md)|Relationship|body|[Body](body.md)|Gets the body object of the cell. Read-only.|1.3|
|[tableCell](tablecell.md)|Relationship|cellPaddingBottom|[float](float.md)|Gets and sets the bottom padding of the cell in points.|1.3|
|[tableCell](tablecell.md)|Relationship|cellPaddingLeft|[float](float.md)|Gets and sets the left padding of the cell in points.|1.3|
|[tableCell](tablecell.md)|Relationship|cellPaddingRight|[float](float.md)|Gets and sets the right padding of the cell in points.|1.3|
|[tableCell](tablecell.md)|Relationship|cellPaddingTop|[float](float.md)|Gets and sets the top padding of the cell in points.|1.3|
|[tableCell](tablecell.md)|Relationship|columnWidth|[float](float.md)|Gets and sets the width of the cell's column in points. This is applicable to uniform tables.|1.3|
|[tableCell](tablecell.md)|Relationship|next|[TableCell](tablecell.md)|Gets the next cell. Read-only.|1.3|
|[tableCell](tablecell.md)|Relationship|parentRow|[TableRow](tablerow.md)|Gets the parent row of the cell. Read-only.|1.3|
|[tableCell](tablecell.md)|Relationship|parentTable|[Table](table.md)|Gets the parent table of the cell. Read-only.|1.3|
|[tableCell](tablecell.md)|Relationship|verticalAlignment|[VerticalAlignment](verticalalignment.md)|Gets and sets the vertical alignment of the cell.|1.3|
|[tableCell](tablecell.md)|Relationship|width|[float](float.md)|Gets the width of the cell in points. Read-only.|1.3|
|[tableCell](tablecell.md)|Method|[deleteColumn()](tablecell.md#deletecolumn)|void|Deletes the column containing this cell. This is applicable to uniform tables.|1.3|
|[tableCell](tablecell.md)|Method|[deleteRow()](tablecell.md#deleterow)|void|Deletes the row containing this cell.|1.3|
|[tableCell](tablecell.md)|Method|[getBorderStyle(borderLocation: BorderLocation)](tablecell.md#getborderstyleborderlocation-borderlocation)|[TableBorderStyle](tableborderstyle.md)|Gets the border style for the specified border.|1.3|
|[tableCell](tablecell.md)|Method|[insertColumns(insertLocation: InsertLocation, columnCount: number, values: string[][])](tablecell.md#insertcolumnsinsertlocation-insertlocation-columncount-number-values-string)|void|Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|1.3|
|[tableCell](tablecell.md)|Method|[insertRows(insertLocation: InsertLocation, rowCount: number, values: string[][])](tablecell.md#insertrowsinsertlocation-insertlocation-rowcount-number-values-string)|void|Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.|1.3|
|[tableCell](tablecell.md)|Method|[split(rowCount: number, columnCount: number)](tablecell.md#splitrowcount-number-columncount-number)|void|Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.|WordApiDesktop, 1.3|
|[tableCellCollection](tablecellcollection.md)|Property|items|[TableCell[]](tablecell.md)|A collection of tableCell objects. Read-only.|1.3|
|[tableCellCollection](tablecellcollection.md)|Relationship|first|[TableCell](tablecell.md)|Gets the first table cell in this collection. Read-only.|1.3|
|[tableCellCollection](tablecellcollection.md)|Method|[getItem(index: number)](tablecellcollection.md#getitemindex-number)|[TableCell](tablecell.md)|Gets a table cell object by its index in the collection.|1.3|
|[tableCollection](tablecollection.md)|Property|items|[Table[]](table.md)|A collection of table objects. Read-only.|1.3|
|[tableCollection](tablecollection.md)|Relationship|first|[Table](table.md)|Gets the first table in this collection. Read-only.|1.3|
|[tableCollection](tablecollection.md)|Method|[getItem(index: number)](tablecollection.md#getitemindex-number)|[Table](table.md)|Gets a table object by its index in the collection.|1.3|
|[tableRow](tablerow.md)|Property|cellCount|int|Gets the number of cells in the row. Read-only.|1.3|
|[tableRow](tablerow.md)|Property|isHeader|bool|Gets a value that indicates whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object. Read-only.|1.3|
|[tableRow](tablerow.md)|Property|rowIndex|int|Gets the index of the row in its parent table. Read-only.|1.3|
|[tableRow](tablerow.md)|Property|shadingColor|string|Gets and sets the shading color.|1.3|
|[tableRow](tablerow.md)|Property|values|string|Gets and sets the text values in the row, as a 1D Javascript array.|1.3|
|[tableRow](tablerow.md)|Relationship|cellPaddingBottom|[float](float.md)|Gets and sets the default bottom cell padding for the row in points.|1.3|
|[tableRow](tablerow.md)|Relationship|cellPaddingLeft|[float](float.md)|Gets and sets the default left cell padding for the row in points.|1.3|
|[tableRow](tablerow.md)|Relationship|cellPaddingRight|[float](float.md)|Gets and sets the default right cell padding for the row in points.|1.3|
|[tableRow](tablerow.md)|Relationship|cellPaddingTop|[float](float.md)|Gets and sets the default top cell padding for the row in points.|1.3|
|[tableRow](tablerow.md)|Relationship|cells|[TableCellCollection](tablecellcollection.md)|Gets cells. Read-only.|1.3|
|[tableRow](tablerow.md)|Relationship|font|[Font](font.md)|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|1.3|
|[tableRow](tablerow.md)|Relationship|next|[TableRow](tablerow.md)|Gets the next row. Read-only.|1.3|
|[tableRow](tablerow.md)|Relationship|parentTable|[Table](table.md)|Gets parent table. Read-only.|1.3|
|[tableRow](tablerow.md)|Relationship|preferredHeight|[float](float.md)|Gets and sets the preferred height of the row in points.|1.3|
|[tableRow](tablerow.md)|Relationship|verticalAlignment|[VerticalAlignment](verticalalignment.md)|Gets and sets the vertical alignment of the cells in the row.|1.3|
|[tableRow](tablerow.md)|Method|[clear()](tablerow.md#clear)|void|Clears the contents of the row.|1.3|
|[tableRow](tablerow.md)|Method|[delete()](tablerow.md#delete)|void|Deletes the entire row.|1.3|
|[tableRow](tablerow.md)|Method|[getBorderStyle(borderLocation: BorderLocation)](tablerow.md#getborderstyleborderlocation-borderlocation)|[TableBorderStyle](tableborderstyle.md)|Gets the border style of the cells in the row.|1.3|
|[tableRow](tablerow.md)|Method|[insertRows(insertLocation: InsertLocation, rowCount: number, values: string[][])](tablerow.md#insertrowsinsertlocation-insertlocation-rowcount-number-values-string)|void|Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.|1.3|
|[tableRow](tablerow.md)|Method|[merge()](tablerow.md#merge)|[TableCell](tablecell.md)|Merges the row into one cell.|WordApiDesktop, 1.3|
|[tableRow](tablerow.md)|Method|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](tablerow.md#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.|1.3|
|[tableRow](tablerow.md)|Method|[select(selectionMode: SelectionMode)](tablerow.md#selectselectionmode-selectionmode)|void|Selects the row and navigates the Word UI to it.|1.3|
|[tableRowCollection](tablerowcollection.md)|Property|items|[TableRow[]](tablerow.md)|A collection of tableRow objects. Read-only.|1.3|
|[tableRowCollection](tablerowcollection.md)|Relationship|first|[TableRow](tablerow.md)|Gets the first row in this collection. Read-only.|1.3|
|[tableRowCollection](tablerowcollection.md)|Method|[getItem(index: number)](tablerowcollection.md#getitemindex-number)|[TableRow](tablerow.md)|Gets a table row object by its index in the collection.|1.3|
