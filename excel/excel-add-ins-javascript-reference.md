## Excel Add-ins JavaScript API reference

_Applies to: Excel 2016, Office 2016_

Below links show the main level Excel Objects and APIs.
	
* [Workbook](resources/workbook.md): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It can be used to also list related references. 
* [Worksheet](resources/worksheet.md): A member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.
	* [Worksheet Collection](resources/worksheetcollection.md): A collection of all the Workbook objects that are part of the workbook. 
* [Range](resources/range.md): Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells.  
* [Table](resources/table.md): Represents a collection of organized cells designed to make management of the data easy. 
	* [Table Collection](resources/tablecollection.md): A collection of Tables in a workbook or worksheet. 
	* [TableColumn Collection](resources/tablecolumncollection.md): A collection of all the columns in a Table. 
	* [TableRow Collection](resources/tablerowcollection.md): A collection of all the rows in a Table. 
* [Chart](resources/chart.md): Represents a Chart object in a worksheet, which is a visual representation of underlying data.   
	* [Chart Collection](resources/chartcollection.md): A collection of charts in a worksheet.	
* [NamedItem](resources/nameditem.md): Represents a defined name for a range of cells or a value. Names can be primitive-named objects, range object, etc.
	* [NamedItem Collection](resources/nameditemcollection.md): A collection of NamedItem objects in a workbook.
* [Binding](resources/binding.md): An abstract class that represents a binding to a section of the workbook.
	* [Binding Collection](resources/bindingcollection.md):A collection of all the Binding objects that are part of the workbook. 
* [TrackedObject Collection](resources/trackedobjectscollection.md): Allows add-ins to add and remove temporary references on range that can be tracked across requests.
* [Request Context](resources/requestcontext.md): The RequestContext object facilitates requests to the Excel application.

##### Learn more

* [Programming overview](excel-add-ins-programming-overview.md): Provides important programming details related to Excel APIs.