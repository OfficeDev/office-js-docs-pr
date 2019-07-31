---
title: Excel JavaScript API requirement set 1.7
description: 'Details about the ExcelApi 1.7 requirement set'
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.7

The Excel JavaScript API requirement set 1.7 features include APIs for charts, events, worksheets, ranges, document properties, named items, protection options and styles.

## Customize charts

With the new chart APIs, you can create additional chart types, add a data series to a chart, set the chart title, add an axis title, add display unit, add a trendline with moving average, change a trendline to linear, and more. The following are some examples:

* Chart axis - get, set, format and remove axis unit, label and title in a chart.
* Chart series - add, set, and delete a series in a chart.  Change series markers, plot orders and sizing.
* Chart trendlines - add, get, and format trendlines in a chart.
* Chart legend - format the legend font in a chart.
* Chart point - set chart point color.
* Chart title substring -  get and set title substring for a chart.
* Chart type - option to create more chart types.

## Events

Excel events APIs provide a variety of event handlers that allow your add-in to automatically run a designated function when a specific event occurs. You can design that function to perform whatever actions your scenario requires. For a list of events that are currently available, see [Work with Events using the Excel JavaScript API](/office/dev/add-ins/excel/excel-add-ins-events).

## Customize the appearance of worksheets and ranges

Using the new APIs, you can customize the appearance of worksheets in multiple ways:

* Freeze panes to keep specific rows or columns visible when you scroll in the worksheet. For example, if the first row in your worksheet contains headers, you might freeze that row so that the column headers will remain visible as you scroll down the worksheet.
* Modify the worksheet tab color.
* Add worksheet headings.

You can customize the appearance of ranges in multiple ways:

* Set the cell style for a range to ensure sure that all cells in the range have consistent formatting. A cell style is a defined set of formatting characteristics, such as fonts and font sizes, number formats, cell borders, and cell shading. Use any of Excel's built-in cell styles or create your own custom cell style.
* Set the text orientation for a range.
* Add or modify a hyperlink on a range that links to another location in the workbook or to an external location.

## Manage document properties

Using the document properties APIs, you can access built-in document properties and also create and manage custom document properties to store state of the workbook and drive workflow and business logic.

## Copy worksheets

Using the worksheet copy APIs, you can copy the data and format from one worksheet to a new worksheet within the same workbook and reduce the amount of data transfer needed.

## Handle ranges with ease

Using the various range APIs, you can do things such as get the surrounding region, get a resized range, and more. These APIs should make tasks like range manipulation and addressing much more efficient.

In addition:

* Workbook and worksheet protection options - use these APIs to protect data in a worksheet and the workbook structure.
* Update a named item - use this API to update a named item.
* Get active cell  - use this API to get the active cell of a workbook.

## API list

| Class | Fields | Description |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Represents the type of the chart. See Excel.ChartType for details.|
||[id](/javascript/api/excel/excel.chart#id)|The unique id of chart. Read-only.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Represents whether to display all field buttons on a PivotChart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#border)|Represents the border format of chart area, which includes color, linestyle, and weight. Read-only.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel.ChartAxisType, group?: Excel.ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Returns the specific axis identified by type and group.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|Returns or sets the base unit for the specified category axis.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Returns or sets the category axis type.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Represents the axis display unit. See Excel.ChartAxisDisplayUnit for details.|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|Represents the base of the logarithm when using logarithmic scales.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Represents the type of major tick mark for the specified axis. See Excel.ChartAxisTickMark for details.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Represents the type of minor tick mark for the specified axis. See Excel.ChartAxisTickMark for details.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Represents the group for the specified axis. See Excel.ChartAxisGroup for details. Read-only.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Represents the custom axis display unit value. Read-only. To set this property, please use the SetCustomDisplayUnit(double) method.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Represents the height, in points, of the chart axis. Null if the axis is not visible. Read-only.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis is not visible. Read-only.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis is not visible. Read-only.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Represents the axis type. See Excel.ChartAxisType for details.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Represents the width, in points, of the chart axis. Null if the axis is not visible. Read-only.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Represents whether Microsoft Excel plots data points from last to first.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Represents the value axis scale type. See Excel.ChartAxisScaleType for details.|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Sets all the category names for the specified axis.|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Sets the axis display unit to a custom value.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Represents whether the axis display unit label is visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Represents the position of tick-mark labels on the specified axis. See Excel.ChartAxisTickLabelPosition for details.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Represents the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Represents the number of categories or series between tick marks.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|A boolean value represents the visibility of the axis.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Clear the border format of a chart element.|
||[color](/javascript/api/excel/excel.chartborder#color)|HTML color code representing the color of borders in the chart.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Represents the line style of the border. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Represents weight of the border, in points.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autotext)|Boolean value representing if data label automatically generates appropriate text based on context.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|String value that represents the formula of chart data label using A1-style notation.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|String value that represents the format code for data label.|
||[position](/javascript/api/excel/excel.chartdatalabel#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Represents the format of chart data label.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible.|
||[separator](/javascript/api/excel/excel.chartdatalabel#separator)|String representing the separator used for the data label on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Boolean value representing if the data label value is visible or not.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|String representing the text of the data label on a chart.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Represents the font attributes, such as font name, font size, color, etc. of chart characters object.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Represents the height, in points, of the legend on the chart. Null if legend is not visible.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Represents the left, in points, of a chart legend. Null if legend is not visible.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Represents a collection of legendEntries in the legend. Read-only.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|Represents if the legend has a shadow on the chart.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Represents the top of a chart legend.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Represents the width, in points, of the legend on the chart. Null if legend is not visible.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Represents the height of the legendEntry on the chart legend.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Represents the index of the legendEntry in the chart legend.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Represents the left of a chart legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Represents the top of a chart legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Represents the width of the legendEntry on the chart Legend.|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Represents the visible of a chart legend entry.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Returns the number of legendEntry in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Returns a legendEntry at the given index.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Gets the loaded child items in this collection.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Represents the line style. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Represents weight of the line, in points.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Represents whether a data point has a data label. Not applicable for surface charts.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Represents marker size of data point.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Represents marker style of a chart data point. See Excel.ChartMarkerStyle for details.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Returns the data label of a chart point. Read-only.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#border)|Represents the border format of a chart data point, which includes color, style, and weight information. Read-only.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Represents the chart type of a series. See Excel.ChartType for details.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Deletes the chart series.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|
||[filtered](/javascript/api/excel/excel.chartseries#filtered)|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Boolean value representing if the series has data labels or not.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Represents markers background color of a chart series.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Represents markers foreground color of a chart series.|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|Represents marker size of a chart series.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|Represents marker style of a chart series. See Excel.ChartMarkerStyle for details.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|Represents the plot order of a chart series within the chart group.|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|Represents a collection of trendlines in the series. Read-only.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Set bubble sizes for a chart series. Only works for bubble charts.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Set values for a chart series. For scatter chart, it means Y axis values.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Set values of X axis for a chart series. Only works for scatter charts.|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|Boolean value representing if the series has a shadow or not.|
||[smooth](/javascript/api/excel/excel.chartseries#smooth)|Boolean value representing if the series is smooth or not. Only applicable to line and scatter charts.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Add a new series to the collection. The new added series is not visible until set values/x axis values/bubble sizes for it (depending on chart type).|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Get the substring of a chart title. Line break '\n' also counts one character.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Represents the horizontal alignment for chart title.|
||[left](/javascript/api/excel/excel.charttitle#left)|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title is not visible.|
||[position](/javascript/api/excel/excel.charttitle#position)|Represents the position of chart title. See Excel.ChartTitlePosition for details.|
||[height](/javascript/api/excel/excel.charttitle#height)|Returns the height, in points, of the chart title. Null if chart title is not visible. Read-only.|
||[width](/javascript/api/excel/excel.charttitle#width)|Returns the width, in points, of the chart title. Null if chart title is not visible. Read-only.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Sets a string value that represents the formula of chart title using A1-style notation.|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|Represents a boolean value that determines if the chart title has a shadow.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttitle#top)|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Represents the vertical alignment of chart title. See Excel.ChartTextVerticalAlignment for details.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#border)|Represents the border format of chart title, which includes color, linestyle, and weight. Read-only.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Represents the number of periods that the trendline extends backward.|
||[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Delete the trendline object.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Represents the number of periods that the trendline extends forward.|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Represents the formatting of a chart trendline.|
||[label](/javascript/api/excel/excel.charttrendline#label)|Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|True if the R-squared for the trendline is displayed on the chart.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Represents the type of a chart trendline.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel.ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Adds a new trendline to trendline collection.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Returns the number of trendlines in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Get trendline object by index, which is the insertion order in items array.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Gets the loaded child items in this collection.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Represents chart line formatting. Read-only.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Deletes the custom property.|
||[key](/javascript/api/excel/excel.customproperty#key)|Gets the key of the custom property. Read only.|
||[type](/javascript/api/excel/excel.customproperty#type)|Gets the value type of the custom property. Read only.|
||[value](/javascript/api/excel/excel.customproperty#value)|Gets or sets the value of the custom property.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Gets the loaded child items in this collection.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Refreshes all the Data Connections in the collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#author)|Gets or sets the author of the workbook.|
||[category](/javascript/api/excel/excel.documentproperties#category)|Gets or sets the category of the workbook.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Gets or sets the comments of the workbook.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Gets or sets the company of the workbook.|
||[keywords](/javascript/api/excel/excel.documentproperties#keywords)|Gets or sets the keywords of the workbook.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|Gets or sets the manager of the workbook.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Gets the creation date of the workbook. Read only.|
||[custom](/javascript/api/excel/excel.documentproperties#custom)|Gets the collection of custom properties of the workbook. Read only.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Gets the last author of the workbook. Read only.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Gets the revision number of the workbook. Read only.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Gets or sets the subject of the workbook.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Gets or sets the title of the workbook.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Returns an object containing values and types of the named item. Read-only.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Represents the types for each item in the named item array|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Represents the values of each item in the named item array.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|
||[getImage()](/javascript/api/excel/excel.range#getimage--)|Renders the range as a base64-encoded png image.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Represents the hyperlink for the current range.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Represents Excel's number format code for the given range as a string in the language of the user.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Represents if the current range is an entire column. Read-only.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Represents if the current range is an entire row. Read-only.|
||[showCard()](/javascript/api/excel/excel.range#showcard--)|Displays the card for an active cell if it has rich value content.|
||[style](/javascript/api/excel/excel.range#style)|Represents the style of the current range.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Gets or sets the text orientation of all the cells within the range.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Determines if the row height of the Range object equals the standard height of the sheet.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Indicates whether the column width of the Range object equals the standard width of the sheet.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Represents the url target for the hyperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Represents the document reference target for the hyperlink.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#screentip)|Represents the string displayed when hovering over the hyperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Represents the string that is displayed in the top left most cell in the range.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[delete()](/javascript/api/excel/excel.style#delete--)|Deletes this style.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Indicates if the formula will be hidden when the worksheet is protected.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Represents the horizontal alignment for the style. See Excel.HorizontalAlignment for details.|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|
||[includeFont](/javascript/api/excel/excel.style#includefont)|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|Indicates if the style includes the NumberFormat property.|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|Indicates if the style includes the FormulaHidden and Locked protection properties.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|An integer from 0 to 250 that indicates the indent level for the style.|
||[locked](/javascript/api/excel/excel.style#locked)|Indicates if the object is locked when the worksheet is protected.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|The format code of the number format for the style.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|The localized format code of the number format for the style.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|The reading order for the style.|
||[borders](/javascript/api/excel/excel.style#borders)|A Border collection of four Border objects that represent the style of the four borders.|
||[builtIn](/javascript/api/excel/excel.style#builtin)|Indicates if the style is a built-in style.|
||[fill](/javascript/api/excel/excel.style#fill)|The Fill of the style.|
||[font](/javascript/api/excel/excel.style#font)|A Font object that represents the font of the style.|
||[name](/javascript/api/excel/excel.style#name)|The name of the style.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|The text orientation for the style.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Represents the vertical alignment for the style. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Indicates if Microsoft Excel wraps the text in the object.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Adds a new style to the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Gets a style by name.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Occurs when data in cells changes on a specific table.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Occurs when the selection changes on a specific table.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Gets the address that represents the changed area of a table on a specific worksheet.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Gets the change type that represents how the Changed event is triggered. See Excel.DataChangeType for details.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Gets the range that represents the changed area of a table on a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Gets the range that represents the changed area of a table on a specific worksheet. It might return null object.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Gets the id of the table in which the data changed.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Occurs when data changes on any table in a workbook, or a worksheet.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Gets the range address that represents the selected area of the table on a specific worksheet.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Indicates if the selection is inside a table, address will be useless if IsInsideTable is false.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Gets the id of the table in which the selection changed.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details. Read-only.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Gets the id of the worksheet in which the selection changed.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Gets the currently active cell from the workbook.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Represents all data connections in the workbook. Read-only.|
||[name](/javascript/api/excel/excel.workbook#name)|Gets the workbook name. Read-only.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Gets the workbook properties. Read-only.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Returns workbook protection object for a workbook. Read-only.|
||[styles](/javascript/api/excel/excel.workbook#styles)|Represents a collection of styles associated with the workbook. Read-only.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protects a workbook. Fails if the workbook has been protected.|
||[protected](/javascript/api/excel/excel.workbookprotection#protected)|Indicates if the workbook is protected. Read-Only.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Unprotects a workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel.WorksheetPositionType, relativeTo?: Excel.Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copy a worksheet and place it at the specified position. Return the copied worksheet.|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Gets an object that can be used to manipulate frozen panes on the worksheet. Read-only.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Occurs when the worksheet is activated.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Occurs when data changed on a specific worksheet.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Occurs when the worksheet is deactivated.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Occurs when the selection changes on a specific worksheet.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|Returns or sets the standard (default) width of all the columns in the worksheet.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|Gets or sets the worksheet tab color.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Gets the id of the worksheet that is activated.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Gets the id of the worksheet that is added to the workbook.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Gets the change type that represents how the Changed event is triggered. See Excel.DataChangeType for details.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Gets the range that represents the changed area of a specific worksheet. It might return null object.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Occurs when any worksheet in the workbook is activated.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Occurs when a new worksheet is added to the workbook.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Occurs when any worksheet in the workbook is deactivated.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Occurs when a worksheet is deleted from the workbook.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Gets the id of the worksheet that is deactivated.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Gets the id of the worksheet that is deleted from the workbook.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Sets the frozen cells in the active worksheet view.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Freeze the first column(s) of the worksheet in place.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Freeze the top row(s) of the worksheet in place.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Gets a range that describes the frozen cells in the active worksheet view.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Gets a range that describes the frozen cells in the active worksheet view.|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Removes all frozen panes in the worksheet.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Unprotects a worksheet.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Represents the worksheet protection option of allowing editing objects.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Represents the worksheet protection option of allowing editing scenarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Represents the worksheet protection option of selection mode.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Gets the range address that represents the selected area of a specific worksheet.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Gets the id of the worksheet in which the selection changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
