---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs'
ms.date: 01/02/2020
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Culture settings | Adjust cultural system settings for the workbook, such as number formatting. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Insert workbook](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insert one workbook into another.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | Save and close workbooks.  | [Workbook](/javascript/api/excel/excel.workbook) |

## API list

The following table lists the Excel JavaScript APIs currently in preview. To see a complete list of all Excel JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Gets the string used as the decimal separator for numeric values. This is based on Excel's local settings.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on Excel's local settings.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Specifies whether the system separators of Microsoft Excel are enabled.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Gets the values from a single dimension of the chart series. These could be either category values or data values, depending on the dimension specified and how the data is mapped for the chart series.|
|[Comment](/javascript/api/excel/excel.comment)|[resolved](/javascript/api/excel/excel.comment#resolved)|Gets or sets the comment thread status. A value of "true" means the comment thread is in the resolved state.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[resolved](/javascript/api/excel/excel.commentreply#resolved)|Gets or sets the comment reply status. A value of "true" means the comment reply is in the resolved state.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Gets the culture name in the format languagecode2-country/regioncode2 (e.g. "zh-cn" or "en-us"). This is based on current system settings.|
||[numberFormatInfo](/javascript/api/excel/excel.cultureinfo#numberformatinfo)|Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Gets the string used as the decimal separator for numeric values. This is based on current system settings.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on current system settings.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies. The returned cell is the intersection of the given row and column that contains the data from the given hierarchy. This method is the inverse of calling getPivotItems and getDataHierarchy on a particular cell.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Gets the range object containing the anchor cell for a cell getting spilled into. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Gets the range object containing the anchor cell for a cell getting spilled into. Read-only.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Gets the range object containing the spill range when called on an anchor cell. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Gets the range object containing the spill range when called on an anchor cell. Read-only.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Represents if all cells have a spill border.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Represents if ALL the cells would be saved as an array formula.|
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
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Returns a collection of worksheet-level custom properties.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Occurs when filter is applied on a specific worksheet.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|The address of the ranges that completed calculation.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Occurs when any worksheet's filter is applied in the workbook.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Gets the key of the custom property. Read only.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Gets the value of the custom property. Read only.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Gets the number of custom properties on this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Gets the loaded child items in this collection.|
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
