# Excel JavaScript API reference

You can use the Excel JavaScript API to build add-ins for Excel 2016. The following list shows the high-level Excel objects that are available in the API. Each object page link contains a description of the properties, relationships, and methods that are available on the object. Explore the links from the menu to learn more.

Note that the relationships section within the document lists the properties that are used to navigate from the main object to another related object. These are non-scalar objects that themselves may contain other properties, methods and relationships.

Some of the core Excel objects are listed below for convenience: 

* [Workbook](workbook.md): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.

* [Worksheet](worksheet.md): Represents a worksheet in a workbook. 
  * [WorksheetCollection](worksheetcollection.md): A collection of the **Worksheet** objects in a workbook.

* [Range](range.md): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.

* [Table](table.md): Represents a collection of organized cells designed to make management of the data easy.
  * [TableCollection](tablecollection.md): A collection of tables in a workbook or worksheet.
  * [TableColumnCollection](tablecolumncollection.md): A collection of all the columns in a table.
  * [TableRowCollection](tablerowcollection.md): A collection of all the rows in a table.

* [Chart](chart.md): Represents a chart object in a worksheet, which is a visual representation of underlying data.
  * [ChartCollection](chartcollection.md): A collection of charts in a worksheet.

* [TableSort](tablesort.md): Represents an object that manages sorting operations on **Table** objects.

* [RangeSort](rangesort.md): Represents a object that manages sorting operations on **Range** objects.

* [Filter](filter.md): Represents an object that manages the filtering of a table's column.

* [WorksheetProtection](worksheetprotection.md): Represents the protection of a **Worksheet** object.

* [NamedItem](nameditem.md): Represents a defined name for a range of cells or a value. 
  * [NamedItemCollection](nameditemcollection.md): A collection of the **NamedItem** objects in a workbook.

* [Binding](binding.md): An abstract class that represents a binding to a section of the workbook.
  * [BindingCollection](bindingcollection.md): A collection of the **Binding** objects in a workbook.

