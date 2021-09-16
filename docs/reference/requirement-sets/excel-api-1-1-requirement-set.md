---
title: Excel JavaScript API requirement set 1.1
description: 'Details about the ExcelApi 1.1 requirement set.'
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
---

# Excel JavaScript API requirement set 1.1

Excel JavaScript API 1.1 is the first version of the API. It is the only Excel-specific requirement set supported by Excel 2016.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.1. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.1, see [Excel APIs in requirement set 1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel.CalculationType)](/javascript/api/excel/excel.application#calculate_calculationType_)|Recalculate all currently opened workbooks in Excel.|
||[calculationMode](/javascript/api/excel/excel.application#calculationMode)|Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getRange__)|Returns the range represented by the binding.|
||[getTable()](/javascript/api/excel/excel.binding#getTable__)|Returns the table represented by the binding.|
||[getText()](/javascript/api/excel/excel.binding#getText__)|Returns the text represented by the binding.|
||[id](/javascript/api/excel/excel.binding#id)|Represents the binding identifier.|
||[type](/javascript/api/excel/excel.binding#type)|Returns the type of the binding.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getItem_id_)|Gets a binding object by ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getItemAt_index_)|Gets a binding object based on its position in the items array.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Returns the number of bindings in the collection.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Gets the loaded child items in this collection.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete__)|Deletes the chart object.|
||[height](/javascript/api/excel/excel.chart#height)|Specifies the height, in points, of the chart object.|
||[left](/javascript/api/excel/excel.chart#left)|The distance, in points, from the left side of the chart to the worksheet origin.|
||[name](/javascript/api/excel/excel.chart#name)|Specifies the name of a chart object.|
||[axes](/javascript/api/excel/excel.chart#axes)|Represents chart axes.|
||[dataLabels](/javascript/api/excel/excel.chart#dataLabels)|Represents the data labels on the chart.|
||[format](/javascript/api/excel/excel.chart#format)|Encapsulates the format properties for the chart area.|
||[legend](/javascript/api/excel/excel.chart#legend)|Represents the legend for the chart.|
||[series](/javascript/api/excel/excel.chart#series)|Represents either a single series or collection of series in the chart.|
||[title](/javascript/api/excel/excel.chart#title)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.|
||[setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chart#setData_sourceData__seriesBy_)|Resets the source data for the chart.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#setPosition_startCell__endCell_)|Positions the chart relative to cells on the worksheet.|
||[top](/javascript/api/excel/excel.chart#top)|Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/javascript/api/excel/excel.chart#width)|Specifies the width, in points, of the chart object.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Represents the font attributes (font name, font size, color, etc.) for the current object.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryAxis)|Represents the category axis in a chart.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesAxis)|Represents the series axis of a 3-D chart.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueAxis)|Represents the value axis in an axis.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorUnit)|Represents the interval between two major tick marks.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Represents the maximum value on the value axis.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Represents the minimum value on the value axis.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorUnit)|Represents the interval between two minor tick marks.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Represents the formatting of a chart object, which includes line and font formatting.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorGridlines)|Returns an object that represents the major gridlines for the specified axis.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorGridlines)|Returns an object that represents the minor gridlines for the specified axis.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Represents the axis title.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Specifies chart line formatting.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Specifies the formatting of the chart axis title.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Specifies the axis title.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Specifies if the axis title is visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add_type__sourceData__seriesBy_)|Creates a new chart.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getItem_name_)|Gets a chart using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getItemAt_index_)|Gets a chart based on its position in the collection.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Returns the number of charts in the worksheet.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Gets the loaded child items in this collection.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Represents the fill format of the current chart data label.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Represents the font attributes (such as font name, font size, and color) for a chart data label.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Value that represents the position of the data label.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Specifies the format of chart data labels, which includes fill and font formatting.|
||[separator](/javascript/api/excel/excel.chartdatalabels#separator)|String representing the separator used for the data labels on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showBubbleSize)|Specifies if the data label bubble size is visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showCategoryName)|Specifies if the data label category name is visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showLegendKey)|Specifies if the data label legend key is visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showPercentage)|Specifies if the data label percentage is visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showSeriesName)|Specifies if the data label series name is visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showValue)|Specifies if the data label value is visible.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear__)|Clears the fill color of a chart element.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setSolidColor_color_)|Sets the fill formatting of a chart element to a uniform color.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.chartfont#color)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.chartfont#name)|Font name (e.g., "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|Size of the font (e.g., 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Type of underline applied to the font.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Represents the formatting of chart gridlines.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Specifies if the axis gridlines are visible.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Represents chart line formatting.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Specifies if the chart legend should overlap with the main body of the chart.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Specifies the position of the legend on the chart.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Represents the formatting of a chart legend, which includes fill and font formatting.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Specifies if the chart legend is visible.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Represents the font attributes such as font name, font size, and color of a chart legend.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear__)|Clears the line format of a chart element.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|HTML color code representing the color of lines in the chart.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Encapsulates the format properties chart point.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Returns the value of a chart point.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Represents the fill format of a chart, which includes background formatting information.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getItemAt_index_)|Retrieve a point based on its position within the series.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Returns the number of chart points in the series.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Gets the loaded child items in this collection.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Specifies the name of a series in a chart.|
||[format](/javascript/api/excel/excel.chartseries#format)|Represents the formatting of a chart series, which includes fill and line formatting.|
||[points](/javascript/api/excel/excel.chartseries#points)|Returns a collection of all points in the series.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getItemAt_index_)|Retrieves a series based on its position in the collection.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Returns the number of series in the collection.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Gets the loaded child items in this collection.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Represents the fill format of a chart series, which includes background formatting information.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Represents line formatting.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|Specifies if the chart title will overlay the chart.|
||[format](/javascript/api/excel/excel.charttitle#format)|Represents the formatting of a chart title, which includes fill and font formatting.|
||[text](/javascript/api/excel/excel.charttitle#text)|Specifies the chart's title text.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Specifies if the chart title is visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Represents the font attributes (such as font name, font size, and color) for an object.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getRange__)|Returns the range object that is associated with the name.|
||[name](/javascript/api/excel/excel.nameditem#name)|The name of the object.|
||[type](/javascript/api/excel/excel.nameditem#type)|Specifies the type of the value returned by the name's formula.|
||[value](/javascript/api/excel/excel.nameditem#value)|Represents the value computed by the name's formula.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Specifies if the object is visible.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getItem_name_)|Gets a `NamedItem` object using its name.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear_applyTo_)|Clear range values, format, fill, border, etc.|
||[delete(shift: Excel.DeleteShiftDirection)](/javascript/api/excel/excel.range#delete_shift_)|Deletes the cells associated with the range.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.range#formulasLocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getBoundingRect_anotherRange_)|Gets the smallest range object that encompasses the given ranges.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getCell_row__column_)|Gets the range object containing the single cell based on row and column numbers.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getColumn_column_)|Gets a column contained in the range.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getEntireColumn__)|Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#getEntireRow__)|Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getIntersection_anotherRange_)|Gets the range object that represents the rectangular intersection of the given ranges.|
||[getLastCell()](/javascript/api/excel/excel.range#getLastCell__)|Gets the last cell within the range.|
||[getLastColumn()](/javascript/api/excel/excel.range#getLastColumn__)|Gets the last column within the range.|
||[getLastRow()](/javascript/api/excel/excel.range#getLastRow__)|Gets the last row within the range.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getOffsetRange_rowOffset__columnOffset_)|Gets an object which represents a range that's offset from the specified range.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getRow_row_)|Gets a row contained in the range.|
||[insert(shift: Excel.InsertShiftDirection)](/javascript/api/excel/excel.range#insert_shift_)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space.|
||[numberFormat](/javascript/api/excel/excel.range#numberFormat)|Represents Excel's number format code for the given range.|
||[address](/javascript/api/excel/excel.range#address)|Specifies the range reference in A1-style.|
||[addressLocal](/javascript/api/excel/excel.range#addressLocal)|Represents the range reference for the specified range in the language of the user.|
||[cellCount](/javascript/api/excel/excel.range#cellCount)|Specifies the number of cells in the range.|
||[columnCount](/javascript/api/excel/excel.range#columnCount)|Specifies the total number of columns in the range.|
||[columnIndex](/javascript/api/excel/excel.range#columnIndex)|Specifies the column number of the first cell in the range.|
||[format](/javascript/api/excel/excel.range#format)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.|
||[rowCount](/javascript/api/excel/excel.range#rowCount)|Returns the total number of rows in the range.|
||[rowIndex](/javascript/api/excel/excel.range#rowIndex)|Returns the row number of the first cell in the range.|
||[text](/javascript/api/excel/excel.range#text)|Text values of the specified range.|
||[valueTypes](/javascript/api/excel/excel.range#valueTypes)|Specifies the type of data in each cell.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|The worksheet containing the current range.|
||[select()](/javascript/api/excel/excel.range#select__)|Selects the specified range in the Excel UI.|
||[values](/javascript/api/excel/excel.range#values)|Represents the raw values of the specified range.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideIndex)|Constant value that indicates the specific side of the border.|
||[style](/javascript/api/excel/excel.rangeborder#style)|One of the constants of line style specifying the line style for the border.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Specifies the weight of the border around a range.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem(index: Excel.BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getItem_index_)|Gets a border object using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getItemAt_index_)|Gets a border object using its index.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Number of border objects in the collection.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Gets the loaded child items in this collection.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear__)|Resets the range background.|
||[color](/javascript/api/excel/excel.rangefill#color)|HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Represents the bold status of the font.|
||[color](/javascript/api/excel/excel.rangefont#color)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Specifies the italic status of the font.|
||[name](/javascript/api/excel/excel.rangefont#name)|Font name (e.g., "Calibri").|
||[size](/javascript/api/excel/excel.rangefont#size)|Font size.|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Type of underline applied to the font.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalAlignment)|Represents the horizontal alignment for the specified object.|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|Collection of border objects that apply to the overall range.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Returns the fill object defined on the overall range.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Returns the font object defined on the overall range.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalAlignment)|Represents the vertical alignment for the specified object.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wrapText)|Specifies if Excel wraps the text in the object.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete__)|Deletes the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getDataBodyRange__)|Gets the range object associated with the data body of the table.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getHeaderRowRange__)|Gets the range object associated with the header row of the table.|
||[getRange()](/javascript/api/excel/excel.table#getRange__)|Gets the range object associated with the entire table.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#getTotalRowRange__)|Gets the range object associated with the totals row of the table.|
||[name](/javascript/api/excel/excel.table#name)|Name of the table.|
||[columns](/javascript/api/excel/excel.table#columns)|Represents a collection of all the columns in the table.|
||[id](/javascript/api/excel/excel.table#id)|Returns a value that uniquely identifies the table in a given workbook.|
||[rows](/javascript/api/excel/excel.table#rows)|Represents a collection of all the rows in the table.|
||[showHeaders](/javascript/api/excel/excel.table#showHeaders)|Specifies if the header row is visible.|
||[showTotals](/javascript/api/excel/excel.table#showTotals)|Specifies if the total row is visible.|
||[style](/javascript/api/excel/excel.table#style)|Constant value that represents the table style.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add_address__hasHeaders_)|Creates a new table.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getItem_key_)|Gets a table by name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getItemAt_index_)|Gets a table based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Returns the number of tables in the workbook.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Gets the loaded child items in this collection.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete__)|Deletes the column from the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getDataBodyRange__)|Gets the range object associated with the data body of the column.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getHeaderRowRange__)|Gets the range object associated with the header row of the column.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getRange__)|Gets the range object associated with the entire column.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#getTotalRowRange__)|Gets the range object associated with the totals row of the column.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Specifies the name of the table column.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Returns a unique key that identifies the column within the table.|
||[index](/javascript/api/excel/excel.tablecolumn#index)|Returns the index number of the column within the columns collection of the table.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Represents the raw values of the specified range.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add_index__values__name_)|Adds a new column to the table.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getItem_key_)|Gets a column object by name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getItemAt_index_)|Gets a column based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Returns the number of columns in the table.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete__)|Deletes the row from the table.|
||[getRange()](/javascript/api/excel/excel.tablerow#getRange__)|Returns the range object associated with the entire row.|
||[index](/javascript/api/excel/excel.tablerow#index)|Returns the index number of the row within the rows collection of the table.|
||[values](/javascript/api/excel/excel.tablerow#values)|Represents the raw values of the specified range.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number)](/javascript/api/excel/excel.tablerowcollection#add_index__values_)|Adds one or more rows to the table.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getItemAt_index_)|Gets a row based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Returns the number of rows in the table.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Gets the loaded child items in this collection.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange()](/javascript/api/excel/excel.workbook#getSelectedRange__)|Gets the currently selected single range from the workbook.|
||[application](/javascript/api/excel/excel.workbook#application)|Represents the Excel application instance that contains this workbook.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Represents a collection of bindings that are part of the workbook.|
||[names](/javascript/api/excel/excel.workbook#names)|Represents a collection of workbook-scoped named items (named ranges and constants).|
||[tables](/javascript/api/excel/excel.workbook#tables)|Represents a collection of tables associated with the workbook.|
||[worksheets](/javascript/api/excel/excel.workbook#worksheets)|Represents a collection of worksheets associated with the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate__)|Activate the worksheet in the Excel UI.|
||[delete()](/javascript/api/excel/excel.worksheet#delete__)|Deletes the worksheet from the workbook.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getCell_row__column_)|Gets the `Range` object containing the single cell based on row and column numbers.|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#getRange_address_)|Gets the `Range` object, representing a single rectangular block of cells, specified by the address or name.|
||[name](/javascript/api/excel/excel.worksheet#name)|The display name of the worksheet.|
||[position](/javascript/api/excel/excel.worksheet#position)|The zero-based position of the worksheet within the workbook.|
||[charts](/javascript/api/excel/excel.worksheet#charts)|Returns a collection of charts that are part of the worksheet.|
||[id](/javascript/api/excel/excel.worksheet#id)|Returns a value that uniquely identifies the worksheet in a given workbook.|
||[tables](/javascript/api/excel/excel.worksheet#tables)|Collection of tables that are part of the worksheet.|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|The visibility of the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#add_name_)|Adds a new worksheet to the workbook.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getActiveWorksheet__)|Gets the currently active worksheet in the workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getItem_key_)|Gets a worksheet object using its name or ID.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Gets the loaded child items in this collection.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
