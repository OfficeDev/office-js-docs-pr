---
title: Excel JavaScript API overview
description: ''
ms.date: 03/19/2019
ms.prod: excel
localization_priority: Priority
---

# Excel JavaScript API overview

You can use the Excel JavaScript API to build add-ins for Excel 2016 or later. The following list shows the high-level Excel objects that are available in the API. Each object page link contains a description of the properties, events, and methods that are available on the object. Explore the links from the menu to learn more.

Some of the core Excel objects are listed below for convenience: 

- [Workbook](/javascript/api/excel/excel.workbook): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.

- [Worksheet](/javascript/api/excel/excel.worksheet): Represents a worksheet in a workbook. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): A collection of the **Worksheet** objects in a workbook.
    - [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): Represents the protection of a **Worksheet** object.

- [Range](/javascript/api/excel/excel.range): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.
    - [ConditionalFormat](/javascript/api/excel/excel.conditionalformat): An object defining a rule and a format applied to the range when the rule's condition is met.
	- [DataValidation](/javascript/api/excel/excel.datavalidation): An object that restricts user input to a range based on a variety of criteria.
    - [RangeSort](/javascript/api/excel/excel.rangesort): Represents a object that manages sorting operations on a range.

- [Table](/javascript/api/excel/excel.table): Represents a collection of organized cells designed to make management of the data easy.
    - [TableCollection](/javascript/api/excel/excel.tablecollection): A collection of tables in a workbook or worksheet.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): A collection of all the columns in a table.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection): A collection of all the rows in a table.
    - [TableSort](/javascript/api/excel/excel.tablesort): Represents an object that manages sorting operations on a table.

- [Chart](/javascript/api/excel/excel.chart): Represents a chart object in a worksheet, which is a visual representation of underlying data.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection): A collection of charts in a worksheet.
	
- [PivotTable](/javascript/api/excel/excel.pivottable): Represents an Excel PivotTable, which is a hierarchical grouping and presentation of data. 
    - [PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): A collection of PivotTables in a worksheet.

- [Filter](/javascript/api/excel/excel.filter): Represents an object that manages the filtering of a table's column.

- [NamedItem](/javascript/api/excel/excel.nameditem): Represents a defined name for a range of cells or a value. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): A collection of the **NamedItem** objects in a workbook.

- [Binding](/javascript/api/excel/excel.binding): An abstract class that represents a binding to a section of the workbook.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection): A collection of the **Binding** objects in a workbook.

## Excel JavaScript API open specifications

As we design and develop new APIs for Excel add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Excel JavaScript APIs, and provide your input on our design specifications.

## Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For detailed information about Excel JavaScript API requirement sets, see the [Excel JavaScript API requirement sets](../requirement-sets/excel-api-requirement-sets.md) article.

## Excel JavaScript API reference

For detailed information about the Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).

## See also

- [Excel add-ins overview](/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office Add-ins platform overview](/office/dev/add-ins/overview/office-add-ins)
- [Excel add-in samples on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
