# Excel JavaScript API reference

You can use the Excel JavaScript API to build add-ins for Excel 2016. The following list shows the high-level Excel objects that are available in the API. Each object page link contains a description of the properties, relationships, and methods that are available on the object. Explore the links to learn more.

* [Workbook](../../reference/excel/workbook.md): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.
* [Worksheet](../../reference/excel/worksheet.md): A member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.
	* [Worksheet Collection](../../reference/excel/worksheetcollection.md): A collection of all the Workbook objects that are part of the workbook.
* [Range](../../reference/excel/range.md): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.
* [Table](../../reference/excel/table.md): Represents a collection of organized cells designed to make management of the data easy.
	* [Table Collection](../../reference/excel/tablecollection.md): A collection of tables in a workbook or worksheet.
	* [TableColumn Collection](../../reference/excel/tablecolumncollection.md): A collection of all the columns in a table.
	* [TableRow Collection](../../reference/excel/tablerowcollection.md): A collection of all the rows in a table.
* [Chart](../../reference/excel/chart.md): Represents a Chart object in a worksheet, which is a visual representation of underlying data.
	* [Chart Collection](../../reference/excel/chartcollection.md): A collection of charts in a worksheet.
* [NamedItem](../../reference/excel/nameditem.md): Represents a defined name for a range of cells or a value. Names can be primitive-named objects, range object, etc.
	* [NamedItem Collection](../../reference/excel/nameditemcollection.md): A collection of NamedItem objects in a workbook.
* [Binding](../../reference/excel/binding.md): An abstract class that represents a binding to a section of the workbook.
	* [Binding Collection](../../reference/excel/bindingcollection.md):A collection of all the Binding objects that are part of the workbook.
* [TrackedObject Collection](../../reference/excel/trackedobjectscollection.md): Allows add-ins to manage a range object reference across sync() batches.
* [Request Context](../../reference/excel/requestcontext.md): The RequestContext object facilitates requests to the Excel application.


##### Additional resources

*  [Excel add-ins programming overview](excel-add-ins-javascript-programming-overview.md)
*  [Build your first Excel add-in](build-your-first-excel-add-in.md)
*  [Snippet Explorer for Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)

