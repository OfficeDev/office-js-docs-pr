---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs'
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Comment Mentions](../../excel/excel-add-ins-comments.md#mentions) | Mention others in comments to send notifications. | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| [Insert Workbook](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insert one workbook into another.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | Save and close workbooks.  | [Workbook](/javascript/api/excel/excel.workbook) |

## API list

The following table lists the Excel JavaScript APIs currently in preview. To see a complete list of all Excel JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview).

| Class | Fields | Description |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Gets the values from a single dimension of the chart series. These could be either category values or data values, depending on the dimension specified and how the data is mapped for the chart series.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Gets the entities (e.g. people) that are mentioned in comments.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Gets the rich comment content (e.g. mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.|
||[resolved](/javascript/api/excel/excel.comment#resolved)|Gets or sets the comment thread status. A value of "true" means the comment thread is in the resolved state.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Gets or sets the email address of the entity that is mentioned in comment.|
||[id](/javascript/api/excel/excel.commentmention#id)|Gets or sets the id of the entity. This is aligned with the id information in `CommentRichContent.richContent`.|
||[name](/javascript/api/excel/excel.commentmention#name)|Gets or sets the name of the entity that is mentioned in comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Gets the entities (e.g. people) that are mentioned in comments.|
||[resolved](/javascript/api/excel/excel.commentreply#resolved)|Gets or sets the comment reply status. A value of "true" means the comment reply is in the resolved state.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Gets the rich comment content (e.g. mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|An array containing all the entities (e.g. people) mentioned within the comment.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies. The returned cell is the intersection of the given row and column that contains the data from the given hierarchy. This method is the inverse of calling getPivotItems and getDataHierarchy on a particular cell.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Gets the range object containing the anchor cell for a cell getting spilled into. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Gets the range object containing the anchor cell for a cell getting spilled into. Read-only.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Gets the range object containing the spill range when called on an anchor cell. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Gets the range object containing the spill range when called on an anchor cell. Read-only.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Represents if all cells have a spill border.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Represents if ALL the cells would be saved as an array formula.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Adjusts the indentation of the range formatting. The indent value ranges from 0 to 250.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet. Returns a Shape object that represents the new image.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Represents the slicer name used in the formula.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Occurs when filter is applied on a specific table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Occurs when filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Represents the id of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Represents the id of the worksheet which contains the table.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Close current workbook.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Save current workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Occurs when filter is applied on a specific worksheet.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Occurs when any worksheet's filter is applied in the workbook.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Represents the id of the worksheet in which the filter is applied.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Gets the type of change that represents how the event was triggered. See `Excel.RowHiddenChangeType` for details.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
