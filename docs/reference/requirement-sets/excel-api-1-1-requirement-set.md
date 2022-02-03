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
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel.CalculationType)](/javascript/api/excel/excel.application#excel-excel-application-calculate-member(1))|Recalculate all currently opened workbooks in Excel.|
||[calculationMode](/javascript/api/excel/excel.application#excel-excel-application-calculationmode-member)|Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#excel-excel-binding-getrange-member(1))|Returns the range represented by the binding.|
||[getTable()](/javascript/api/excel/excel.binding#excel-excel-binding-gettable-member(1))|Returns the table represented by the binding.|
||[getText()](/javascript/api/excel/excel.binding#excel-excel-binding-gettext-member(1))|Returns the text represented by the binding.|
||[id](/javascript/api/excel/excel.binding#excel-excel-binding-id-member)|Represents the binding identifier.|
||[type](/javascript/api/excel/excel.binding#excel-excel-binding-type-member)|Returns the type of the binding.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[count](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-count-member)|Returns the number of bindings in the collection.|
||[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitem-member(1))|Gets a binding object by ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemat-member(1))|Gets a binding object based on its position in the items array.|
||[items](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-items-member)|Gets the loaded child items in this collection.|
|[Chart](/javascript/api/excel/excel.chart)|[axes](/javascript/api/excel/excel.chart#excel-excel-chart-axes-member)|Represents chart axes.|
||[dataLabels](/javascript/api/excel/excel.chart#excel-excel-chart-datalabels-member)|Represents the data labels on the chart.|
||[delete()](/javascript/api/excel/excel.chart#excel-excel-chart-delete-member(1))|Deletes the chart object.|
||[format](/javascript/api/excel/excel.chart#excel-excel-chart-format-member)|Encapsulates the format properties for the chart area.|
||[height](/javascript/api/excel/excel.chart#excel-excel-chart-height-member)|Specifies the height, in points, of the chart object.|
||[left](/javascript/api/excel/excel.chart#excel-excel-chart-left-member)|The distance, in points, from the left side of the chart to the worksheet origin.|
||[legend](/javascript/api/excel/excel.chart#excel-excel-chart-legend-member)|Represents the legend for the chart.|
||[name](/javascript/api/excel/excel.chart#excel-excel-chart-name-member)|Specifies the name of a chart object.|
||[series](/javascript/api/excel/excel.chart#excel-excel-chart-series-member)|Represents either a single series or collection of series in the chart.|
||[setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chart#excel-excel-chart-setdata-member(1))|Resets the source data for the chart.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#excel-excel-chart-setposition-member(1))|Positions the chart relative to cells on the worksheet.|
||[title](/javascript/api/excel/excel.chart#excel-excel-chart-title-member)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.|
||[top](/javascript/api/excel/excel.chart#excel-excel-chart-top-member)|Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/javascript/api/excel/excel.chart#excel-excel-chart-width-member)|Specifies the width, in points, of the chart object.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-font-member)|Represents the font attributes (font name, font size, color, etc.) for the current object.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-categoryaxis-member)|Represents the category axis in a chart.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-seriesaxis-member)|Represents the series axis of a 3-D chart.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-valueaxis-member)|Represents the value axis in an axis.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[format](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-format-member)|Represents the formatting of a chart object, which includes line and font formatting.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorgridlines-member)|Returns an object that represents the major gridlines for the specified axis.|
||[majorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorunit-member)|Represents the interval between two major tick marks.|
||[maximum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-maximum-member)|Represents the maximum value on the value axis.|
||[minimum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minimum-member)|Represents the minimum value on the value axis.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorgridlines-member)|Returns an object that represents the minor gridlines for the specified axis.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorunit-member)|Represents the interval between two minor tick marks.|
||[title](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-title-member)|Represents the axis title.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-font-member)|Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.|
||[line](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-line-member)|Specifies chart line formatting.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-format-member)|Specifies the formatting of the chart axis title.|
||[text](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-text-member)|Specifies the axis title.|
||[visible](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-visible-member)|Specifies if the axis title is visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-font-member)|Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-add-member(1))|Creates a new chart.|
||[count](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-count-member)|Returns the number of charts in the worksheet.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitem-member(1))|Gets a chart using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemat-member(1))|Gets a chart based on its position in the collection.|
||[items](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-items-member)|Gets the loaded child items in this collection.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-fill-member)|Represents the fill format of the current chart data label.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-font-member)|Represents the font attributes (such as font name, font size, and color) for a chart data label.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[format](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-format-member)|Specifies the format of chart data labels, which includes fill and font formatting.|
||[position](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-position-member)|Value that represents the position of the data label.|
||[separator](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-separator-member)|String representing the separator used for the data labels on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showbubblesize-member)|Specifies if the data label bubble size is visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showcategoryname-member)|Specifies if the data label category name is visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showlegendkey-member)|Specifies if the data label legend key is visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showpercentage-member)|Specifies if the data label percentage is visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showseriesname-member)|Specifies if the data label series name is visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showvalue-member)|Specifies if the data label value is visible.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-clear-member(1))|Clears the fill color of a chart element.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-setsolidcolor-member(1))|Sets the fill formatting of a chart element to a uniform color.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-bold-member)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-color-member)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-italic-member)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-name-member)|Font name (e.g., "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-size-member)|Size of the font (e.g., 11)|
||[underline](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-underline-member)|Type of underline applied to the font.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-format-member)|Represents the formatting of chart gridlines.|
||[visible](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-visible-member)|Specifies if the axis gridlines are visible.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#excel-excel-chartgridlinesformat-line-member)|Represents chart line formatting.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[format](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-format-member)|Represents the formatting of a chart legend, which includes fill and font formatting.|
||[overlay](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-overlay-member)|Specifies if the chart legend should overlap with the main body of the chart.|
||[position](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-position-member)|Specifies the position of the legend on the chart.|
||[visible](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-visible-member)|Specifies if the chart legend is visible.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-font-member)|Represents the font attributes such as font name, font size, and color of a chart legend.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-clear-member(1))|Clears the line format of a chart element.|
||[color](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-color-member)|HTML color code representing the color of lines in the chart.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-format-member)|Encapsulates the format properties chart point.|
||[value](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-value-member)|Returns the value of a chart point.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-fill-member)|Represents the fill format of a chart, which includes background formatting information.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[count](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-count-member)|Returns the number of chart points in the series.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getitemat-member(1))|Retrieve a point based on its position within the series.|
||[items](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-items-member)|Gets the loaded child items in this collection.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[format](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-format-member)|Represents the formatting of a chart series, which includes fill and line formatting.|
||[name](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-name-member)|Specifies the name of a series in a chart.|
||[points](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-points-member)|Returns a collection of all points in the series.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[count](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-count-member)|Returns the number of series in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getitemat-member(1))|Retrieves a series based on its position in the collection.|
||[items](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-items-member)|Gets the loaded child items in this collection.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-fill-member)|Represents the fill format of a chart series, which includes background formatting information.|
||[line](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-line-member)|Represents line formatting.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[format](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-format-member)|Represents the formatting of a chart title, which includes fill and font formatting.|
||[overlay](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-overlay-member)|Specifies if the chart title will overlay the chart.|
||[text](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-text-member)|Specifies the chart's title text.|
||[visible](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-visible-member)|Specifies if the chart title is visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-font-member)|Represents the font attributes (such as font name, font size, and color) for an object.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrange-member(1))|Returns the range object that is associated with the name.|
||[name](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-name-member)|The name of the object.|
||[type](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-type-member)|Specifies the type of the value returned by the name's formula.|
||[value](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-value-member)|Represents the value computed by the name's formula.|
||[visible](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-visible-member)|Specifies if the object is visible.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitem-member(1))|Gets a `NamedItem` object using its name.|
||[items](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[address](/javascript/api/excel/excel.range#excel-excel-range-address-member)|Specifies the range reference in A1-style.|
||[addressLocal](/javascript/api/excel/excel.range#excel-excel-range-addresslocal-member)|Represents the range reference for the specified range in the language of the user.|
||[cellCount](/javascript/api/excel/excel.range#excel-excel-range-cellcount-member)|Specifies the number of cells in the range.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#excel-excel-range-clear-member(1))|Clear range values, format, fill, border, etc.|
||[columnCount](/javascript/api/excel/excel.range#excel-excel-range-columncount-member)|Specifies the total number of columns in the range.|
||[columnIndex](/javascript/api/excel/excel.range#excel-excel-range-columnindex-member)|Specifies the column number of the first cell in the range.|
||[delete(shift: Excel.DeleteShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-delete-member(1))|Deletes the cells associated with the range.|
||[format](/javascript/api/excel/excel.range#excel-excel-range-format-member)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.|
||[formulas](/javascript/api/excel/excel.range#excel-excel-range-formulas-member)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.range#excel-excel-range-formulaslocal-member)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getboundingrect-member(1))|Gets the smallest range object that encompasses the given ranges.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcell-member(1))|Gets the range object containing the single cell based on row and column numbers.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcolumn-member(1))|Gets a column contained in the range.|
||[getEntireColumn()](/javascript/api/excel/excel.range#excel-excel-range-getentirecolumn-member(1))|Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#excel-excel-range-getentirerow-member(1))|Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersection-member(1))|Gets the range object that represents the rectangular intersection of the given ranges.|
||[getLastCell()](/javascript/api/excel/excel.range#excel-excel-range-getlastcell-member(1))|Gets the last cell within the range.|
||[getLastColumn()](/javascript/api/excel/excel.range#excel-excel-range-getlastcolumn-member(1))|Gets the last column within the range.|
||[getLastRow()](/javascript/api/excel/excel.range#excel-excel-range-getlastrow-member(1))|Gets the last row within the range.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1))|Gets an object which represents a range that's offset from the specified range.|
||[getRow(row: number)](/javascript/api/excel/excel.range#excel-excel-range-getrow-member(1))|Gets a row contained in the range.|
||[insert(shift: Excel.InsertShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1))|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space.|
||[numberFormat](/javascript/api/excel/excel.range#excel-excel-range-numberformat-member)|Represents Excel's number format code for the given range.|
||[rowCount](/javascript/api/excel/excel.range#excel-excel-range-rowcount-member)|Returns the total number of rows in the range.|
||[rowIndex](/javascript/api/excel/excel.range#excel-excel-range-rowindex-member)|Returns the row number of the first cell in the range.|
||[select()](/javascript/api/excel/excel.range#excel-excel-range-select-member(1))|Selects the specified range in the Excel UI.|
||[text](/javascript/api/excel/excel.range#excel-excel-range-text-member)|Text values of the specified range.|
||[valueTypes](/javascript/api/excel/excel.range#excel-excel-range-valuetypes-member)|Specifies the type of data in each cell.|
||[values](/javascript/api/excel/excel.range#excel-excel-range-values-member)|Represents the raw values of the specified range.|
||[worksheet](/javascript/api/excel/excel.range#excel-excel-range-worksheet-member)|The worksheet containing the current range.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-color-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-sideindex-member)|Constant value that indicates the specific side of the border.|
||[style](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-style-member)|One of the constants of line style specifying the line style for the border.|
||[weight](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-weight-member)|Specifies the weight of the border around a range.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[count](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-count-member)|Number of border objects in the collection.|
||[getItem(index: Excel.BorderIndex)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitem-member(1))|Gets a border object using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitemat-member(1))|Gets a border object using its index.|
||[items](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-items-member)|Gets the loaded child items in this collection.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-clear-member(1))|Resets the range background.|
||[color](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-color-member)|HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-bold-member)|Represents the bold status of the font.|
||[color](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-color-member)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-italic-member)|Specifies the italic status of the font.|
||[name](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-name-member)|Font name (e.g., "Calibri").|
||[size](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-size-member)|Font size.|
||[underline](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-underline-member)|Type of underline applied to the font.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[borders](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-borders-member)|Collection of border objects that apply to the overall range.|
||[fill](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-fill-member)|Returns the fill object defined on the overall range.|
||[font](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-font-member)|Returns the font object defined on the overall range.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-horizontalalignment-member)|Represents the horizontal alignment for the specified object.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-verticalalignment-member)|Represents the vertical alignment for the specified object.|
||[wrapText](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-wraptext-member)|Specifies if Excel wraps the text in the object.|
|[Table](/javascript/api/excel/excel.table)|[columns](/javascript/api/excel/excel.table#excel-excel-table-columns-member)|Represents a collection of all the columns in the table.|
||[delete()](/javascript/api/excel/excel.table#excel-excel-table-delete-member(1))|Deletes the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#excel-excel-table-getdatabodyrange-member(1))|Gets the range object associated with the data body of the table.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#excel-excel-table-getheaderrowrange-member(1))|Gets the range object associated with the header row of the table.|
||[getRange()](/javascript/api/excel/excel.table#excel-excel-table-getrange-member(1))|Gets the range object associated with the entire table.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#excel-excel-table-gettotalrowrange-member(1))|Gets the range object associated with the totals row of the table.|
||[id](/javascript/api/excel/excel.table#excel-excel-table-id-member)|Returns a value that uniquely identifies the table in a given workbook.|
||[name](/javascript/api/excel/excel.table#excel-excel-table-name-member)|Name of the table.|
||[rows](/javascript/api/excel/excel.table#excel-excel-table-rows-member)|Represents a collection of all the rows in the table.|
||[showHeaders](/javascript/api/excel/excel.table#excel-excel-table-showheaders-member)|Specifies if the header row is visible.|
||[showTotals](/javascript/api/excel/excel.table#excel-excel-table-showtotals-member)|Specifies if the total row is visible.|
||[style](/javascript/api/excel/excel.table#excel-excel-table-style-member)|Constant value that represents the table style.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-add-member(1))|Creates a new table.|
||[count](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-count-member)|Returns the number of tables in the workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitem-member(1))|Gets a table by name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemat-member(1))|Gets a table based on its position in the collection.|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|Gets the loaded child items in this collection.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-delete-member(1))|Deletes the column from the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getdatabodyrange-member(1))|Gets the range object associated with the data body of the column.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getheaderrowrange-member(1))|Gets the range object associated with the header row of the column.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getrange-member(1))|Gets the range object associated with the entire column.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-gettotalrowrange-member(1))|Gets the range object associated with the totals row of the column.|
||[id](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-id-member)|Returns a unique key that identifies the column within the table.|
||[index](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-index-member)|Returns the index number of the column within the columns collection of the table.|
||[name](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-name-member)|Specifies the name of the table column.|
||[values](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-values-member)|Represents the raw values of the specified range.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-add-member(1))|Adds a new column to the table.|
||[count](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-count-member)|Returns the number of columns in the table.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitem-member(1))|Gets a column object by name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemat-member(1))|Gets a column based on its position in the collection.|
||[items](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-delete-member(1))|Deletes the row from the table.|
||[getRange()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-getrange-member(1))|Returns the range object associated with the entire row.|
||[index](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-index-member)|Returns the index number of the row within the rows collection of the table.|
||[values](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-values-member)|Represents the raw values of the specified range.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, alwaysInsert?: boolean)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))|Adds one or more rows to the table.|
||[count](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-count-member)|Returns the number of rows in the table.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getitemat-member(1))|Gets a row based on its position in the collection.|
||[items](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-items-member)|Gets the loaded child items in this collection.|
|[Workbook](/javascript/api/excel/excel.workbook)|[application](/javascript/api/excel/excel.workbook#excel-excel-workbook-application-member)|Represents the Excel application instance that contains this workbook.|
||[bindings](/javascript/api/excel/excel.workbook#excel-excel-workbook-bindings-member)|Represents a collection of bindings that are part of the workbook.|
||[getSelectedRange()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1))|Gets the currently selected single range from the workbook.|
||[names](/javascript/api/excel/excel.workbook#excel-excel-workbook-names-member)|Represents a collection of workbook-scoped named items (named ranges and constants).|
||[tables](/javascript/api/excel/excel.workbook#excel-excel-workbook-tables-member)|Represents a collection of tables associated with the workbook.|
||[worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member)|Represents a collection of worksheets associated with the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-activate-member(1))|Activate the worksheet in the Excel UI.|
||[charts](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-charts-member)|Returns a collection of charts that are part of the worksheet.|
||[delete()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-delete-member(1))|Deletes the worksheet from the workbook.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getcell-member(1))|Gets the `Range` object containing the single cell based on row and column numbers.|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1))|Gets the `Range` object, representing a single rectangular block of cells, specified by the address or name.|
||[id](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-id-member)|Returns a value that uniquely identifies the worksheet in a given workbook.|
||[name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member)|The display name of the worksheet.|
||[position](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-position-member)|The zero-based position of the worksheet within the workbook.|
||[tables](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tables-member)|Collection of tables that are part of the worksheet.|
||[visibility](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-visibility-member)|The visibility of the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-add-member(1))|Adds a new worksheet to the workbook.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getactiveworksheet-member(1))|Gets the currently active worksheet in the workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitem-member(1))|Gets a worksheet object using its name or ID.|
||[items](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-items-member)|Gets the loaded child items in this collection.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
