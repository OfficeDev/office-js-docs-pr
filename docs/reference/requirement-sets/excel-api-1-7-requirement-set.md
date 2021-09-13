---
title: Excel JavaScript API requirement set 1.7
description: 'Details about the ExcelApi 1.7 requirement set.'
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.7

The Excel JavaScript API requirement set 1.7 features include APIs for charts, events, worksheets, ranges, document properties, named items, protection options and styles.

## Customize charts

With the new chart APIs, you can create additional chart types, add a data series to a chart, set the chart title, add an axis title, add display unit, add a trendline with moving average, change a trendline to linear, and more. The following are some examples.

- Chart axis - get, set, format and remove axis unit, label and title in a chart.
- Chart series - add, set, and delete a series in a chart.  Change series markers, plot orders and sizing.
- Chart trendlines - add, get, and format trendlines in a chart.
- Chart legend - format the legend font in a chart.
- Chart point - set chart point color.
- Chart title substring -  get and set title substring for a chart.
- Chart type - option to create more chart types.

## Events

Excel events APIs provide a variety of event handlers that allow your add-in to automatically run a designated function when a specific event occurs. You can design that function to perform whatever actions your scenario requires. For a list of events that are currently available, see [Work with Events using the Excel JavaScript API](../../excel/excel-add-ins-events.md).

## Customize the appearance of worksheets and ranges

Using the new APIs, you can customize the appearance of worksheets in multiple ways:

- Freeze panes to keep specific rows or columns visible when you scroll in the worksheet. For example, if the first row in your worksheet contains headers, you might freeze that row so that the column headers will remain visible as you scroll down the worksheet.
- Modify the worksheet tab color.
- Add worksheet headings.

You can customize the appearance of ranges in multiple ways:

- Set the cell style for a range to ensure sure that all cells in the range have consistent formatting. A cell style is a defined set of formatting characteristics, such as fonts and font sizes, number formats, cell borders, and cell shading. Use any of Excel's built-in cell styles or create your own custom cell style.
- Set the text orientation for a range.
- Add or modify a hyperlink on a range that links to another location in the workbook or to an external location.

## Manage document properties

Using the document properties APIs, you can access built-in document properties and also create and manage custom document properties to store state of the workbook and drive workflow and business logic.

## Copy worksheets

Using the worksheet copy APIs, you can copy the data and format from one worksheet to a new worksheet within the same workbook and reduce the amount of data transfer needed.

## Handle ranges with ease

Using the various range APIs, you can do things such as get the surrounding region, get a resized range, and more. These APIs should make tasks like range manipulation and addressing much more efficient.

In addition:

- Workbook and worksheet protection options - use these APIs to protect data in a worksheet and the workbook structure.
- Update a named item - use this API to update a named item.
- Get active cell  - use this API to get the active cell of a workbook.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.7. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.7 or earlier, see [Excel APIs in requirement set 1.7 or earlier](/javascript/api/excel?view=excel-js-1.7&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#chartType)|Specifies the type of the chart.|
||[id](/javascript/api/excel/excel.chart#id)|The unique ID of chart.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showAllFieldButtons)|Specifies whether to display all field buttons on a PivotChart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#border)|Represents the border format of chart area, which includes color, linestyle, and weight.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel.ChartAxisType, group?: Excel.ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getItem_type__group_)|Returns the specific axis identified by type and group.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#baseTimeUnit)|Specifies the base unit for the specified category axis.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categoryType)|Specifies the category axis type.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayUnit)|Represents the axis display unit.|
||[logBase](/javascript/api/excel/excel.chartaxis#logBase)|Specifies the base of the logarithm when using logarithmic scales.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majorTickMark)|Specifies the type of major tick mark for the specified axis.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majorTimeUnitScale)|Specifies the major unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minorTickMark)|Specifies the type of minor tick mark for the specified axis.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minorTimeUnitScale)|Specifies the minor unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisGroup)|Specifies the group for the specified axis.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customDisplayUnit)|Specifies the custom axis display unit value.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Specifies the height, in points, of the chart axis.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Specifies the distance, in points, from the left edge of the axis to the left of chart area.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Specifies the distance, in points, from the top edge of the axis to the top of chart area.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Specifies the axis type.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Specifies the width, in points, of the chart axis.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reversePlotOrder)|Specifies if Excel plots data points from last to first.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaleType)|Specifies the value axis scale type.|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#setCategoryNames_sourceData_)|Sets all the category names for the specified axis.|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#setCustomDisplayUnit_value_)|Sets the axis display unit to a custom value.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showDisplayUnitLabel)|Specifies if the axis display unit label is visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#tickLabelPosition)|Specifies the position of tick-mark labels on the specified axis.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#tickLabelSpacing)|Specifies the number of categories or series between tick-mark labels.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickMarkSpacing)|Specifies the number of categories or series between tick marks.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Specifies if the axis is visible.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|HTML color code representing the color of borders in the chart.|
||[lineStyle](/javascript/api/excel/excel.chartborder#lineStyle)|Represents the line style of the border.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Represents weight of the border, in points.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Value that represents the position of the data label.|
||[separator](/javascript/api/excel/excel.chartdatalabel#separator)|String representing the separator used for the data label on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showBubbleSize)|Specifies if the data label bubble size is visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showCategoryName)|Specifies if the data label category name is visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showLegendKey)|Specifies if the data label legend key is visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showPercentage)|Specifies if the data label percentage is visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showSeriesName)|Specifies if the data label series name is visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showValue)|Specifies if the data label value is visible.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Represents the font attributes, such as font name, font size, and color of a chart characters object.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Specifies the height, in points, of the legend on the chart.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Specifies the left value, in points, of the legend on the chart.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendEntries)|Represents a collection of legendEntries in the legend.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showShadow)|Specifies if the legend has a shadow on the chart.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Specifies the top of a chart legend.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Specifies the width, in points, of the legend on the chart.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Represents the visibility of a chart legend entry.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getCount__)|Returns the number of legend entries in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getItemAt_index_)|Returns a legend entry at the given index.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Gets the loaded child items in this collection.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#lineStyle)|Represents the line style.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Represents weight of the line, in points.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasDataLabel)|Represents whether a data point has a data label.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerBackgroundColor)|HTML color code representation of the marker background color of a data point (e.g., #FF0000 represents Red).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerForegroundColor)|HTML color code representation of the marker foreground color of a data point (e.g., #FF0000 represents Red).|
||[markerSize](/javascript/api/excel/excel.chartpoint#markerSize)|Represents marker size of a data point.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerStyle)|Represents marker style of a chart data point.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#dataLabel)|Returns the data label of a chart point.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#border)|Represents the border format of a chart data point, which includes color, style, and weight information.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#chartType)|Represents the chart type of a series.|
||[delete()](/javascript/api/excel/excel.chartseries#delete__)|Deletes the chart series.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutHoleSize)|Represents the doughnut hole size of a chart series.|
||[filtered](/javascript/api/excel/excel.chartseries#filtered)|Specifies if the series is filtered.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapWidth)|Represents the gap width of a chart series.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasDataLabels)|Specifies if the series has data labels.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerBackgroundColor)|Specifies the marker background color of a chart series.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerForegroundColor)|Specifies the marker foreground color of a chart series.|
||[markerSize](/javascript/api/excel/excel.chartseries#markerSize)|Specifies the marker size of a chart series.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerStyle)|Specifies the marker style of a chart series.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotOrder)|Specifies the plot order of a chart series within the chart group.|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|The collection of trendlines in the series.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#setBubbleSizes_sourceData_)|Sets the bubble sizes for a chart series.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setValues_sourceData_)|Sets the values for a chart series.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setXAxisValues_sourceData_)|Sets the values of the x-axis for a chart series.|
||[showShadow](/javascript/api/excel/excel.chartseries#showShadow)|Specifies if the series has a shadow.|
||[smooth](/javascript/api/excel/excel.chartseries#smooth)|Specifies if the series is smooth.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#add_name__index_)|Add a new series to the collection.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#getSubstring_start__length_)|Get the substring of a chart title.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalAlignment)|Specifies the horizontal alignment for chart title.|
||[left](/javascript/api/excel/excel.charttitle#left)|Specifies the distance, in points, from the left edge of chart title to the left edge of chart area.|
||[position](/javascript/api/excel/excel.charttitle#position)|Represents the position of chart title.|
||[height](/javascript/api/excel/excel.charttitle#height)|Returns the height, in points, of the chart title.|
||[width](/javascript/api/excel/excel.charttitle#width)|Specifies the width, in points, of the chart title.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#setFormula_formula_)|Sets a string value that represents the formula of chart title using A1-style notation.|
||[showShadow](/javascript/api/excel/excel.charttitle#showShadow)|Represents a boolean value that determines if the chart title has a shadow.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textOrientation)|Specifies the angle to which the text is oriented for the chart title.|
||[top](/javascript/api/excel/excel.charttitle#top)|Specifies the distance, in points, from the top edge of chart title to the top of chart area.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalAlignment)|Specifies the vertical alignment of chart title.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#border)|Represents the border format of chart title, which includes color, linestyle, and weight.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete__)|Delete the trendline object.|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|Represents the intercept value of the trendline.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingAveragePeriod)|Represents the period of a chart trendline.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Represents the name of the trendline.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialOrder)|Represents the order of a chart trendline.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Represents the formatting of a chart trendline.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Represents the type of a chart trendline.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel.ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add_type_)|Adds a new trendline to trendline collection.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getCount__)|Returns the number of trendlines in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getItem_index_)|Gets a trendline object by index, which is the insertion order in the items array.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Gets the loaded child items in this collection.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Represents chart line formatting.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete__)|Deletes the custom property.|
||[key](/javascript/api/excel/excel.customproperty#key)|The key of the custom property.|
||[type](/javascript/api/excel/excel.customproperty#type)|The type of the value used for the custom property.|
||[value](/javascript/api/excel/excel.customproperty#value)|The value of the custom property.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add_key__value_)|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#deleteAll__)|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getCount__)|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getItem_key_)|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getItemOrNullObject_key_)|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Gets the loaded child items in this collection.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#refreshAll__)|Refreshes all the data connections in the collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#author)|The author of the workbook.|
||[category](/javascript/api/excel/excel.documentproperties#category)|The category of the workbook.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|The comments of the workbook.|
||[company](/javascript/api/excel/excel.documentproperties#company)|The company of the workbook.|
||[keywords](/javascript/api/excel/excel.documentproperties#keywords)|The keywords of the workbook.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|The manager of the workbook.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationDate)|Gets the creation date of the workbook.|
||[custom](/javascript/api/excel/excel.documentproperties#custom)|Gets the collection of custom properties of the workbook.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastAuthor)|Gets the last author of the workbook.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionNumber)|Gets the revision number of the workbook.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|The subject of the workbook.|
||[title](/javascript/api/excel/excel.documentproperties#title)|The title of the workbook.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|The formula of the named item.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayValues)|Returns an object containing values and types of the named item.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Represents the types for each item in the named item array|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Represents the values of each item in the named item array.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#getAbsoluteResizedRange_numRows__numColumns_)|Gets a `Range` object with the same top-left cell as the current `Range` object, but with the specified numbers of rows and columns.|
||[getImage()](/javascript/api/excel/excel.range#getImage__)|Renders the range as a base64-encoded png image.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getSurroundingRegion__)|Returns a `Range` object that represents the surrounding region for the top-left cell in this range.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Represents the hyperlink for the current range.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberFormatLocal)|Represents Excel's number format code for the given range, based on the language settings of the user.|
||[isEntireColumn](/javascript/api/excel/excel.range#isEntireColumn)|Represents if the current range is an entire column.|
||[isEntireRow](/javascript/api/excel/excel.range#isEntireRow)|Represents if the current range is an entire row.|
||[showCard()](/javascript/api/excel/excel.range#showCard__)|Displays the card for an active cell if it has rich value content.|
||[style](/javascript/api/excel/excel.range#style)|Represents the style of the current range.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textOrientation)|The text orientation of all the cells within the range.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#useStandardHeight)|Determines if the row height of the `Range` object equals the standard height of the sheet.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#useStandardWidth)|Specifies if the column width of the `Range` object equals the standard width of the sheet.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Represents the URL target for the hyperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentReference)|Represents the document reference target for the hyperlink.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#screenTip)|Represents the string displayed when hovering over the hyperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#textToDisplay)|Represents the string that is displayed in the top left most cell in the range.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete__)|Deletes this style.|
||[formulaHidden](/javascript/api/excel/excel.style#formulaHidden)|Specifies if the formula will be hidden when the worksheet is protected.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalAlignment)|Represents the horizontal alignment for the style.|
||[includeAlignment](/javascript/api/excel/excel.style#includeAlignment)|Specifies if the style includes the auto indent, horizontal alignment, vertical alignment, wrap text, indent level, and text orientation properties.|
||[includeBorder](/javascript/api/excel/excel.style#includeBorder)|Specifies if the style includes the color, color index, line style, and weight border properties.|
||[includeFont](/javascript/api/excel/excel.style#includeFont)|Specifies if the style includes the background, bold, color, color index, font style, italic, name, size, strikethrough, subscript, superscript, and underline font properties.|
||[includeNumber](/javascript/api/excel/excel.style#includeNumber)|Specifies if the style includes the number format property.|
||[includePatterns](/javascript/api/excel/excel.style#includePatterns)|Specifies if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.|
||[includeProtection](/javascript/api/excel/excel.style#includeProtection)|Specifies if the style includes the formula hidden and locked protection properties.|
||[indentLevel](/javascript/api/excel/excel.style#indentLevel)|An integer from 0 to 250 that indicates the indent level for the style.|
||[locked](/javascript/api/excel/excel.style#locked)|Specifies if the object is locked when the worksheet is protected.|
||[numberFormat](/javascript/api/excel/excel.style#numberFormat)|The format code of the number format for the style.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberFormatLocal)|The localized format code of the number format for the style.|
||[readingOrder](/javascript/api/excel/excel.style#readingOrder)|The reading order for the style.|
||[borders](/javascript/api/excel/excel.style#borders)|A collection of four border objects that represent the style of the four borders.|
||[builtIn](/javascript/api/excel/excel.style#builtIn)|Specifies if the style is a built-in style.|
||[fill](/javascript/api/excel/excel.style#fill)|The fill of the style.|
||[font](/javascript/api/excel/excel.style#font)|A `Font` object that represents the font of the style.|
||[name](/javascript/api/excel/excel.style#name)|The name of the style.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinkToFit)|Specifies if text automatically shrinks to fit in the available column width.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalAlignment)|Specifies the vertical alignment for the style.|
||[wrapText](/javascript/api/excel/excel.style#wrapText)|Specifies if Excel wraps the text in the object.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add_name_)|Adds a new style to the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getItem_name_)|Gets a `Style` by name.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onChanged)|Occurs when data in cells changes on a specific table.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onSelectionChanged)|Occurs when the selection changes on a specific table.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Gets the address that represents the changed area of a table on a specific worksheet.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changeType)|Gets the change type that represents how the changed event is triggered.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Gets the source of the event.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableId)|Gets the ID of the table in which the data changed.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetId)|Gets the ID of the worksheet in which the data changed.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onChanged)|Occurs when data changes on any table in a workbook, or a worksheet.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Gets the range address that represents the selected area of the table on a specific worksheet.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isInsideTable)|Specifies if the selection is inside a table.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableId)|Gets the ID of the table in which the selection changed.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetId)|Gets the ID of the worksheet in which the selection changed.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getActiveCell__)|Gets the currently active cell from the workbook.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataConnections)|Represents all data connections in the workbook.|
||[name](/javascript/api/excel/excel.workbook#name)|Gets the workbook name.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Gets the workbook properties.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Returns the protection object for a workbook.|
||[styles](/javascript/api/excel/excel.workbook#styles)|Represents a collection of styles associated with the workbook.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#protect_password_)|Protects a workbook.|
||[protected](/javascript/api/excel/excel.workbookprotection#protected)|Specifies if the workbook is protected.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect_password_)|Unprotects a workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel.WorksheetPositionType, relativeTo?: Excel.Worksheet)](/javascript/api/excel/excel.worksheet#copy_positionType__relativeTo_)|Copies a worksheet and places it at the specified position.|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getRangeByIndexes_startRow__startColumn__rowCount__columnCount_)|Gets the `Range` object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezePanes)|Gets an object that can be used to manipulate frozen panes on the worksheet.|
||[onActivated](/javascript/api/excel/excel.worksheet#onActivated)|Occurs when the worksheet is activated.|
||[onChanged](/javascript/api/excel/excel.worksheet#onChanged)|Occurs when data changes in a specific worksheet.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#onDeactivated)|Occurs when the worksheet is deactivated.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onSelectionChanged)|Occurs when the selection changes on a specific worksheet.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardHeight)|Returns the standard (default) height of all the rows in the worksheet, in points.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardWidth)|Specifies the standard (default) width of all the columns in the worksheet.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabColor)|The tab color of the worksheet.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetId)|Gets the ID of the worksheet that is activated.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetId)|Gets the ID of the worksheet that is added to the workbook.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changeType)|Gets the change type that represents how the changed event is triggered.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetId)|Gets the ID of the worksheet in which the data changed.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onActivated)|Occurs when any worksheet in the workbook is activated.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onAdded)|Occurs when a new worksheet is added to the workbook.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#onDeactivated)|Occurs when any worksheet in the workbook is deactivated.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#onDeleted)|Occurs when a worksheet is deleted from the workbook.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetId)|Gets the ID of the worksheet that is deactivated.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetId)|Gets the ID of the worksheet that is deleted from the workbook.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#freezeAt_frozenRange_)|Sets the frozen cells in the active worksheet view.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezeColumns_count_)|Freeze the first column or columns of the worksheet in place.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezeRows_count_)|Freeze the top row or rows of the worksheet in place.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getLocation__)|Gets a range that describes the frozen cells in the active worksheet view.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getLocationOrNullObject__)|Gets a range that describes the frozen cells in the active worksheet view.|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze__)|Removes all frozen panes in the worksheet.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect_password_)|Unprotects a worksheet.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#allowEditObjects)|Represents the worksheet protection option allowing editing of objects.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#allowEditScenarios)|Represents the worksheet protection option allowing editing of scenarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionMode)|Represents the worksheet protection option of selection mode.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Gets the range address that represents the selected area of a specific worksheet.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetId)|Gets the ID of the worksheet in which the selection changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
