---
title: Excel JavaScript API requirement set 1.1
description: 'Details about the ExcelApi 1.1 requirement set'
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript API requirement set 1.1

Excel JavaScript API 1.1 is the first version of the API. It is the only Excel-specific requirement set supported by Excel 2016.

## API list

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel.CalculationType)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalculate all currently opened workbooks in Excel.|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|Returns the calculation mode used in the workbook, as defined by the constants in Excel.CalculationMode. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Returns the range represented by the binding. Will throw an error if binding is not of the correct type.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Returns the table represented by the binding. Will throw an error if binding is not of the correct type.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Returns the text represented by the binding. Will throw an error if binding is not of the correct type.|
||[id](/javascript/api/excel/excel.binding#id)|Represents binding identifier. Read-only.|
||[type](/javascript/api/excel/excel.binding#type)|Returns the type of the binding. See Excel.BindingType for details. Read-only.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Gets a binding object by ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Gets a binding object based on its position in the items array.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Returns the number of bindings in the collection. Read-only.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Gets the loaded child items in this collection.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Deletes the chart object.|
||[height](/javascript/api/excel/excel.chart#height)|Represents the height, in points, of the chart object.|
||[left](/javascript/api/excel/excel.chart#left)|The distance, in points, from the left side of the chart to the worksheet origin.|
||[name](/javascript/api/excel/excel.chart#name)|Represents the name of a chart object.|
||[axes](/javascript/api/excel/excel.chart#axes)|Represents chart axes. Read-only.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Represents the datalabels on the chart. Read-only.|
||[format](/javascript/api/excel/excel.chart#format)|Encapsulates the format properties for the chart area. Read-only.|
||[legend](/javascript/api/excel/excel.chart#legend)|Represents the legend for the chart. Read-only.|
||[series](/javascript/api/excel/excel.chart#series)|Represents either a single series or collection of series in the chart. Read-only.|
||[title](/javascript/api/excel/excel.chart#title)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title. Read-only.|
||[setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Resets the source data for the chart.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Positions the chart relative to cells on the worksheet.|
||[top](/javascript/api/excel/excel.chart#top)|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/javascript/api/excel/excel.chart#width)|Represents the width, in points, of the chart object.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Represents the fill format of an object, which includes background formatting information. Read-only.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|Represents the category axis in a chart. Read-only.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|Represents the series axis of a 3-dimensional chart. Read-only.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Represents the value axis in an axis. Read-only.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Represents the formatting of a chart object, which includes line and font formatting. Read-only.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Returns a Gridlines object that represents the major gridlines for the specified axis. Read-only.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Represents the axis title. Read-only.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Represents chart line formatting. Read-only.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Represents the formatting of chart axis title. Read-only.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Represents the axis title.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|A boolean that specifies the visibility of an axis title.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Creates a new chart.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Gets a chart based on its position in the collection.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Returns the number of charts in the worksheet. Read-only.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Gets the loaded child items in this collection.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Represents the fill format of the current chart data label. Read-only.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Represents the format of chart data labels, which includes fill and font formatting. Read-only.|
||[separator](/javascript/api/excel/excel.chartdatalabels#separator)|String representing the separator used for the data labels on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Clear the fill color of a chart element.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Sets the fill formatting of a chart element to a uniform color.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.chartfont#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.chartfont#name)|Font name (e.g. "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|Size of the font (e.g. 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Represents the formatting of chart gridlines. Read-only.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Boolean value representing if the axis gridlines are visible or not.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Represents chart line formatting. Read-only.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Boolean value for whether the chart legend should overlap with the main body of the chart.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|A boolean value the represents the visibility of a ChartLegend object.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Represents the fill format of an object, which includes background formatting information. Read-only.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Clear the line format of a chart element.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|HTML color code representing the color of lines in the chart.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Encapsulates the format properties chart point. Read-only.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Returns the value of a chart point. Read-only.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Represents the fill format of a chart, which includes background formatting information. Read-only.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Retrieve a point based on its position within the series.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Returns the number of chart points in the series. Read-only.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Gets the loaded child items in this collection.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Represents the name of a series in a chart.|
||[format](/javascript/api/excel/excel.chartseries#format)|Represents the formatting of a chart series, which includes fill and line formatting. Read-only.|
||[points](/javascript/api/excel/excel.chartseries#points)|Represents a collection of all points in the series. Read-only.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Retrieves a series based on its position in the collection.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Returns the number of series in the collection. Read-only.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Gets the loaded child items in this collection.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Represents the fill format of a chart series, which includes background formatting information. Read-only.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Represents line formatting. Read-only.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|Boolean value representing if the chart title will overlay the chart or not.|
||[format](/javascript/api/excel/excel.charttitle#format)|Represents the formatting of a chart title, which includes fill and font formatting. Read-only.|
||[text](/javascript/api/excel/excel.charttitle#text)|Represents the title text of a chart.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|A boolean value the represents the visibility of a chart title object.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Represents the fill format of an object, which includes background formatting information. Read-only.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.|
||[name](/javascript/api/excel/excel.nameditem#name)|The name of the object. Read-only.|
||[type](/javascript/api/excel/excel.nameditem#type)|Indicates the type of the value returned by the name's formula. See Excel.NamedItemType for details. Read-only.|
||[value](/javascript/api/excel/excel.nameditem#value)|Represents the value computed by the name's formula. For a named range, will return the range address. Read-only.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Specifies whether the object is visible or not.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Gets a NamedItem object using its name.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Clear range values, format, fill, border, etc.|
||[delete(shift: Excel.DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|Deletes the cells associated with the range.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E15".|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Gets a column contained in the range.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Gets the range object that represents the rectangular intersection of the given ranges.|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".|
||[getLastColumn()](/javascript/api/excel/excel.range#getlastcolumn--)|Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".|
||[getLastRow()](/javascript/api/excel/excel.range#getlastrow--)|Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Gets a row contained in the range.|
||[insert(shift: Excel.InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Represents Excel's number format code for the given range.|
||[address](/javascript/api/excel/excel.range#address)|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. "Sheet1!A1:B4"). Read-only.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Represents range reference for the specified range in the language of the user. Read-only.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|
||[columnCount](/javascript/api/excel/excel.range#columncount)|Represents the total number of columns in the range. Read-only.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|
||[format](/javascript/api/excel/excel.range#format)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.|
||[rowCount](/javascript/api/excel/excel.range#rowcount)|Returns the total number of rows in the range. Read-only.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|
||[text](/javascript/api/excel/excel.range#text)|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Represents the type of data of each cell. Read-only.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|The worksheet containing the current range. Read-only.|
||[select()](/javascript/api/excel/excel.range#select--)|Selects the specified range in the Excel UI.|
||[values](/javascript/api/excel/excel.range#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.|
||[style](/javascript/api/excel/excel.rangeborder#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Specifies the weight of the border around a range. See Excel.BorderWeight for details.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem(index: Excel.BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Gets a border object using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Gets a border object using its index.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Number of border objects in the collection. Read-only.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Gets the loaded child items in this collection.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Resets the range background.|
||[color](/javascript/api/excel/excel.rangefill#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.rangefont#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.rangefont#name)|Font name (e.g. "Calibri")|
||[size](/javascript/api/excel/excel.rangefont#size)|Font size.|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|Collection of border objects that apply to the overall range. Read-only.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Returns the fill object defined on the overall range. Read-only.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Returns the font object defined on the overall range. Read-only.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Deletes the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|Gets the range object associated with the data body of the table.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|Gets the range object associated with header row of the table.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Gets the range object associated with the entire table.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|Gets the range object associated with totals row of the table.|
||[name](/javascript/api/excel/excel.table#name)|Name of the table.|
||[columns](/javascript/api/excel/excel.table#columns)|Represents a collection of all the columns in the table. Read-only.|
||[id](/javascript/api/excel/excel.table#id)|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
||[rows](/javascript/api/excel/excel.table#rows)|Represents a collection of all the rows in the table. Read-only.|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
||[style](/javascript/api/excel/excel.table#style)|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Gets a table by Name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Gets a table based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Returns the number of tables in the workbook. Read-only.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Gets the loaded child items in this collection.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Deletes the column from the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Gets the range object associated with the data body of the column.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Gets the range object associated with the header row of the column.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Gets the range object associated with the entire column.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Gets the range object associated with the totals row of the column.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Represents the name of the table column.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Returns a unique key that identifies the column within the table. Read-only.|
||[index](/javascript/api/excel/excel.tablecolumn#index)|Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Adds a new column to the table.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Gets a column object by Name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Gets a column based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Returns the number of columns in the table. Read-only.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Deletes the row from the table.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Returns the range object associated with the entire row.|
||[index](/javascript/api/excel/excel.tablerow#index)|Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.|
||[values](/javascript/api/excel/excel.tablerow#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Adds one or more rows to the table. The return object will be the top of the newly added row(s).|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Gets a row based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Returns the number of rows in the table. Read-only.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Gets the loaded child items in this collection.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange()](/javascript/api/excel/excel.workbook#getselectedrange--)|Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.|
||[application](/javascript/api/excel/excel.workbook#application)|Represents the Excel application instance that contains this workbook. Read-only.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Represents a collection of bindings that are part of the workbook. Read-only.|
||[names](/javascript/api/excel/excel.workbook#names)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.|
||[tables](/javascript/api/excel/excel.workbook#tables)|Represents a collection of tables associated with the workbook. Read-only.|
||[worksheets](/javascript/api/excel/excel.workbook#worksheets)|Represents a collection of worksheets associated with the workbook. Read-only.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Activate the worksheet in the Excel UI.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Deletes the worksheet from the workbook. Note that if the worksheet's visibility is set to "VeryHidden", the delete operation will fail with a GeneralException.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#getrange-address-)|Gets the range object, representing a single rectangular block of cells, specified by the address or name.|
||[name](/javascript/api/excel/excel.worksheet#name)|The display name of the worksheet.|
||[position](/javascript/api/excel/excel.worksheet#position)|The zero-based position of the worksheet within the workbook.|
||[charts](/javascript/api/excel/excel.worksheet#charts)|Returns collection of charts that are part of the worksheet. Read-only.|
||[id](/javascript/api/excel/excel.worksheet#id)|Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|
||[tables](/javascript/api/excel/excel.worksheet#tables)|Collection of tables that are part of the worksheet. Read-only.|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|The Visibility of the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Gets the currently active worksheet in the workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Gets a worksheet object using its Name or ID.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Gets the loaded child items in this collection.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
