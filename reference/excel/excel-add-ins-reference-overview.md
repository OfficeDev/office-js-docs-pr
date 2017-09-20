# Excel JavaScript API reference

You can use the Excel JavaScript API to build add-ins for Excel 2016. The following list shows the high-level Excel objects that are available in the API. Each object page link contains a description of the properties, relationships, and methods that are available on the object. Explore the links from the menu to learn more.

Note that the relationships section within the document lists the properties that are used to navigate from main object to another related object. These are non-scalar objects that themselves may contain other properties, methods and relationships.

Some of the core Excel objects are listed below for convenience: 

* [Workbook](workbook.md): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.
* [Worksheet](worksheet.md): A member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.
  * [Worksheet Collection](worksheetcollection.md): A collection of all the Worksheet objects that are part of the workbook.
* [Range](range.md): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.
* [Table](table.md): Represents a collection of organized cells designed to make management of the data easy.
  * [Table Collection](tablecollection.md): A collection of tables in a workbook or worksheet.
  * [TableColumn Collection](tablecolumncollection.md): A collection of all the columns in a table.
  * [TableRow Collection](tablerowcollection.md): A collection of all the rows in a table.
* [Chart](chart.md): Represents a Chart object in a worksheet, which is a visual representation of underlying data.
  * [Chart Collection](chartcollection.md): A collection of charts in a worksheet.
* [TableSort](tablesort.md): Represents a object that sorting operations on Table objects.
* [RangeSort](rangesort.md): Represents a object that sorting operations on Range objects.
* [Filter](filter.md): Represents a fitler object that manages the filtering of a table's column.
* [Worksheet Protection](worksheetprotection.md): Represents the protection of a worksheet object.
* [NamedItem](nameditem.md): Represents a defined name for a range of cells or a value. Names can be primitive-named objects, range object, etc.
  * [NamedItem Collection](nameditemcollection.md): A collection of NamedItem objects in a workbook.
* [Binding](binding.md): An abstract class that represents a binding to a section of the workbook.
  * [Binding Collection](bindingcollection.md):A collection of all the Binding objects that are part of the workbook.

