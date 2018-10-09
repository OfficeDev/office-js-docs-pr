# Excel JavaScript API overview

You can use the Excel JavaScript API to build add-ins for Excel 2016 or later. The following list shows the high-level Excel objects that are available in the API. Each object page link contains a description of the properties, events, and methods that are available on the object. Explore the links from the menu to learn more.

Some of the core Excel objects are listed below for convenience: 

- [Workbook](/javascript/api/excel/excel.workbook): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.

- [Worksheet](/javascript/api/excel/excel.worksheet): Represents a worksheet in a workbook. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): A collection of the **Worksheet** objects in a workbook.

- [Range](/javascript/api/excel/excel.range): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.

- [Table](/javascript/api/excel/excel.table): Represents a collection of organized cells designed to make management of the data easy.
    - [TableCollection](/javascript/api/excel/excel.tablecollection): A collection of tables in a workbook or worksheet.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): A collection of all the columns in a table.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection): A collection of all the rows in a table.

- [Chart](/javascript/api/excel/excel.chart): Represents a chart object in a worksheet, which is a visual representation of underlying data.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection): A collection of charts in a worksheet.

- [TableSort](/javascript/api/excel/excel.tablesort): Represents an object that manages sorting operations on **Table** objects.

- [RangeSort](/javascript/api/excel/excel.rangesort): Represents a object that manages sorting operations on **Range** objects.

- [Filter](/javascript/api/excel/excel.filter): Represents an object that manages the filtering of a table's column.

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): Represents the protection of a **Worksheet** object.

- [NamedItem](/javascript/api/excel/excel.nameditem): Represents a defined name for a range of cells or a value. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): A collection of the **NamedItem** objects in a workbook.

- [Binding](/javascript/api/excel/excel.binding): An abstract class that represents a binding to a section of the workbook.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection): A collection of the **Binding** objects in a workbook.

## Excel JavaScript API open specifications

As we design and develop new APIs for Excel add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Excel JavaScript APIs, and provide your input on our design specifications.

## Excel JavaScript API reference

For detailed information about Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).

## See also

- [Excel add-ins overview](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office Add-ins platform overview](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Excel add-in samples on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
