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
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#excel-excel-chart-charttype-member)|Specifies the type of the chart.|
||[id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member)|The unique ID of chart.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#excel-excel-chart-showallfieldbuttons-member)|Specifies whether to display all field buttons on a PivotChart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-border-member)|Represents the border format of chart area, which includes color, linestyle, and weight.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel.ChartAxisType, group?: Excel.ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-getitem-member(1))|Returns the specific axis identified by type and group.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[axisGroup](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-axisgroup-member)|Specifies the group for the specified axis.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-basetimeunit-member)|Specifies the base unit for the specified category axis.|
||[categoryType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-categorytype-member)|Specifies the category axis type.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-customdisplayunit-member)|Specifies the custom axis display unit value.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-displayunit-member)|Represents the axis display unit.|
||[height](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-height-member)|Specifies the height, in points, of the chart axis.|
||[left](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-left-member)|Specifies the distance, in points, from the left edge of the axis to the left of chart area.|
||[logBase](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-logbase-member)|Specifies the base of the logarithm when using logarithmic scales.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortickmark-member)|Specifies the type of major tick mark for the specified axis.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortimeunitscale-member)|Specifies the major unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortickmark-member)|Specifies the type of minor tick mark for the specified axis.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortimeunitscale-member)|Specifies the minor unit scale value for the category axis when the `categoryType` property is set to `dateAxis`.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-reverseplotorder-member)|Specifies if Excel plots data points from last to first.|
||[scaleType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-scaletype-member)|Specifies the value axis scale type.|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcategorynames-member(1))|Sets all the category names for the specified axis.|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcustomdisplayunit-member(1))|Sets the axis display unit to a custom value.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-showdisplayunitlabel-member)|Specifies if the axis display unit label is visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelposition-member)|Specifies the position of tick-mark labels on the specified axis.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelspacing-member)|Specifies the number of categories or series between tick-mark labels.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-tickmarkspacing-member)|Specifies the number of categories or series between tick marks.|
||[top](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-top-member)|Specifies the distance, in points, from the top edge of the axis to the top of chart area.|
||[type](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-type-member)|Specifies the axis type.|
||[visible](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-visible-member)|Specifies if the axis is visible.|
||[width](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-width-member)|Specifies the width, in points, of the chart axis.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-color-member)|HTML color code representing the color of borders in the chart.|
||[lineStyle](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-linestyle-member)|Represents the line style of the border.|
||[weight](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-weight-member)|Represents weight of the border, in points.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-position-member)|Value that represents the position of the data label.|
||[separator](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-separator-member)|String representing the separator used for the data label on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showbubblesize-member)|Specifies if the data label bubble size is visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showcategoryname-member)|Specifies if the data label category name is visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showlegendkey-member)|Specifies if the data label legend key is visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showpercentage-member)|Specifies if the data label percentage is visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showseriesname-member)|Specifies if the data label series name is visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showvalue-member)|Specifies if the data label value is visible.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#excel-excel-chartformatstring-font-member)|Represents the font attributes, such as font name, font size, and color of a chart characters object.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-height-member)|Specifies the height, in points, of the legend on the chart.|
||[left](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-left-member)|Specifies the left value, in points, of the legend on the chart.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-legendentries-member)|Represents a collection of legendEntries in the legend.|
||[showShadow](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-showshadow-member)|Specifies if the legend has a shadow on the chart.|
||[top](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-top-member)|Specifies the top of a chart legend.|
||[width](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-width-member)|Specifies the width, in points, of the legend on the chart.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-visible-member)|Represents the visibility of a chart legend entry.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getcount-member(1))|Returns the number of legend entries in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getitemat-member(1))|Returns a legend entry at the given index.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-items-member)|Gets the loaded child items in this collection.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-linestyle-member)|Represents the line style.|
||[weight](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-weight-member)|Represents weight of the line, in points.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[dataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-datalabel-member)|Returns the data label of a chart point.|
||[hasDataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-hasdatalabel-member)|Represents whether a data point has a data label.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerbackgroundcolor-member)|HTML color code representation of the marker background color of a data point (e.g., #FF0000 represents Red).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerforegroundcolor-member)|HTML color code representation of the marker foreground color of a data point (e.g., #FF0000 represents Red).|
||[markerSize](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markersize-member)|Represents marker size of a data point.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerstyle-member)|Represents marker style of a chart data point.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-border-member)|Represents the border format of a chart data point, which includes color, style, and weight information.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-charttype-member)|Represents the chart type of a series.|
||[delete()](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-delete-member(1))|Deletes the chart series.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-doughnutholesize-member)|Represents the doughnut hole size of a chart series.|
||[filtered](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-filtered-member)|Specifies if the series is filtered.|
||[gapWidth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gapwidth-member)|Represents the gap width of a chart series.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-hasdatalabels-member)|Specifies if the series has data labels.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerbackgroundcolor-member)|Specifies the marker background color of a chart series.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerforegroundcolor-member)|Specifies the marker foreground color of a chart series.|
||[markerSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markersize-member)|Specifies the marker size of a chart series.|
||[markerStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerstyle-member)|Specifies the marker style of a chart series.|
||[plotOrder](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-plotorder-member)|Specifies the plot order of a chart series within the chart group.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setbubblesizes-member(1))|Sets the bubble sizes for a chart series.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setvalues-member(1))|Sets the values for a chart series.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setxaxisvalues-member(1))|Sets the values of the x-axis for a chart series.|
||[showShadow](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showshadow-member)|Specifies if the series has a shadow.|
||[smooth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-smooth-member)|Specifies if the series is smooth.|
||[trendlines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-trendlines-member)|The collection of trendlines in the series.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-add-member(1))|Add a new series to the collection.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-getsubstring-member(1))|Get the substring of a chart title.|
||[height](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-height-member)|Returns the height, in points, of the chart title.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-horizontalalignment-member)|Specifies the horizontal alignment for chart title.|
||[left](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-left-member)|Specifies the distance, in points, from the left edge of chart title to the left edge of chart area.|
||[position](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-position-member)|Represents the position of chart title.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-setformula-member(1))|Sets a string value that represents the formula of chart title using A1-style notation.|
||[showShadow](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-showshadow-member)|Represents a boolean value that determines if the chart title has a shadow.|
||[textOrientation](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-textorientation-member)|Specifies the angle to which the text is oriented for the chart title.|
||[top](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-top-member)|Specifies the distance, in points, from the top edge of chart title to the top of chart area.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-verticalalignment-member)|Specifies the vertical alignment of chart title.|
||[width](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-width-member)|Specifies the width, in points, of the chart title.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-border-member)|Represents the border format of chart title, which includes color, linestyle, and weight.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-delete-member(1))|Delete the trendline object.|
||[format](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-format-member)|Represents the formatting of a chart trendline.|
||[intercept](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-intercept-member)|Represents the intercept value of the trendline.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-movingaverageperiod-member)|Represents the period of a chart trendline.|
||[name](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-name-member)|Represents the name of the trendline.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-polynomialorder-member)|Represents the order of a chart trendline.|
||[type](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-type-member)|Represents the type of a chart trendline.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel.ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-add-member(1))|Adds a new trendline to trendline collection.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getcount-member(1))|Returns the number of trendlines in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getitem-member(1))|Gets a trendline object by index, which is the insertion order in the items array.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-items-member)|Gets the loaded child items in this collection.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#excel-excel-charttrendlineformat-line-member)|Represents chart line formatting.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-delete-member(1))|Deletes the custom property.|
||[key](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-key-member)|The key of the custom property.|
||[type](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-type-member)|The type of the value used for the custom property.|
||[value](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-value-member)|The value of the custom property.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-add-member(1))|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-deleteall-member(1))|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getcount-member(1))|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitemornullobject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-items-member)|Gets the loaded child items in this collection.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#excel-excel-dataconnectioncollection-refreshall-member(1))|Refreshes all the data connections in the collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-author-member)|The author of the workbook.|
||[category](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-category-member)|The category of the workbook.|
||[comments](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-comments-member)|The comments of the workbook.|
||[company](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-company-member)|The company of the workbook.|
||[creationDate](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-creationdate-member)|Gets the creation date of the workbook.|
||[custom](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member)|Gets the collection of custom properties of the workbook.|
||[keywords](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-keywords-member)|The keywords of the workbook.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-lastauthor-member)|Gets the last author of the workbook.|
||[manager](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-manager-member)|The manager of the workbook.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-revisionnumber-member)|Gets the revision number of the workbook.|
||[subject](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-subject-member)|The subject of the workbook.|
||[title](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-title-member)|The title of the workbook.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[arrayValues](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-arrayvalues-member)|Returns an object containing values and types of the named item.|
||[formula](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-formula-member)|The formula of the named item.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-types-member)|Represents the types for each item in the named item array|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-values-member)|Represents the values of each item in the named item array.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#excel-excel-range-getabsoluteresizedrange-member(1))|Gets a `Range` object with the same top-left cell as the current `Range` object, but with the specified numbers of rows and columns.|
||[getImage()](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1))|Renders the range as a base64-encoded png image.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#excel-excel-range-getsurroundingregion-member(1))|Returns a `Range` object that represents the surrounding region for the top-left cell in this range.|
||[hyperlink](/javascript/api/excel/excel.range#excel-excel-range-hyperlink-member)|Represents the hyperlink for the current range.|
||[isEntireColumn](/javascript/api/excel/excel.range#excel-excel-range-isentirecolumn-member)|Represents if the current range is an entire column.|
||[isEntireRow](/javascript/api/excel/excel.range#excel-excel-range-isentirerow-member)|Represents if the current range is an entire row.|
||[numberFormatLocal](/javascript/api/excel/excel.range#excel-excel-range-numberformatlocal-member)|Represents Excel's number format code for the given range, based on the language settings of the user.|
||[showCard()](/javascript/api/excel/excel.range#excel-excel-range-showcard-member(1))|Displays the card for an active cell if it has rich value content.|
||[style](/javascript/api/excel/excel.range#excel-excel-range-style-member)|Represents the style of the current range.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-textorientation-member)|The text orientation of all the cells within the range.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardheight-member)|Determines if the row height of the `Range` object equals the standard height of the sheet.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardwidth-member)|Specifies if the column width of the `Range` object equals the standard width of the sheet.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-address-member)|Represents the URL target for the hyperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-documentreference-member)|Represents the document reference target for the hyperlink.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-screentip-member)|Represents the string displayed when hovering over the hyperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-texttodisplay-member)|Represents the string that is displayed in the top left most cell in the range.|
|[Style](/javascript/api/excel/excel.style)|[borders](/javascript/api/excel/excel.style#excel-excel-style-borders-member)|A collection of four border objects that represent the style of the four borders.|
||[builtIn](/javascript/api/excel/excel.style#excel-excel-style-builtin-member)|Specifies if the style is a built-in style.|
||[delete()](/javascript/api/excel/excel.style#excel-excel-style-delete-member(1))|Deletes this style.|
||[fill](/javascript/api/excel/excel.style#excel-excel-style-fill-member)|The fill of the style.|
||[font](/javascript/api/excel/excel.style#excel-excel-style-font-member)|A `Font` object that represents the font of the style.|
||[formulaHidden](/javascript/api/excel/excel.style#excel-excel-style-formulahidden-member)|Specifies if the formula will be hidden when the worksheet is protected.|
||[horizontalAlignment](/javascript/api/excel/excel.style#excel-excel-style-horizontalalignment-member)|Represents the horizontal alignment for the style.|
||[includeAlignment](/javascript/api/excel/excel.style#excel-excel-style-includealignment-member)|Specifies if the style includes the auto indent, horizontal alignment, vertical alignment, wrap text, indent level, and text orientation properties.|
||[includeBorder](/javascript/api/excel/excel.style#excel-excel-style-includeborder-member)|Specifies if the style includes the color, color index, line style, and weight border properties.|
||[includeFont](/javascript/api/excel/excel.style#excel-excel-style-includefont-member)|Specifies if the style includes the background, bold, color, color index, font style, italic, name, size, strikethrough, subscript, superscript, and underline font properties.|
||[includeNumber](/javascript/api/excel/excel.style#excel-excel-style-includenumber-member)|Specifies if the style includes the number format property.|
||[includePatterns](/javascript/api/excel/excel.style#excel-excel-style-includepatterns-member)|Specifies if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.|
||[includeProtection](/javascript/api/excel/excel.style#excel-excel-style-includeprotection-member)|Specifies if the style includes the formula hidden and locked protection properties.|
||[indentLevel](/javascript/api/excel/excel.style#excel-excel-style-indentlevel-member)|An integer from 0 to 250 that indicates the indent level for the style.|
||[locked](/javascript/api/excel/excel.style#excel-excel-style-locked-member)|Specifies if the object is locked when the worksheet is protected.|
||[name](/javascript/api/excel/excel.style#excel-excel-style-name-member)|The name of the style.|
||[numberFormat](/javascript/api/excel/excel.style#excel-excel-style-numberformat-member)|The format code of the number format for the style.|
||[numberFormatLocal](/javascript/api/excel/excel.style#excel-excel-style-numberformatlocal-member)|The localized format code of the number format for the style.|
||[readingOrder](/javascript/api/excel/excel.style#excel-excel-style-readingorder-member)|The reading order for the style.|
||[shrinkToFit](/javascript/api/excel/excel.style#excel-excel-style-shrinktofit-member)|Specifies if text automatically shrinks to fit in the available column width.|
||[verticalAlignment](/javascript/api/excel/excel.style#excel-excel-style-verticalalignment-member)|Specifies the vertical alignment for the style.|
||[wrapText](/javascript/api/excel/excel.style#excel-excel-style-wraptext-member)|Specifies if Excel wraps the text in the object.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-add-member(1))|Adds a new style to the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitem-member(1))|Gets a `Style` by name.|
||[items](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member)|Occurs when data in cells changes on a specific table.|
||[onSelectionChanged](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)|Occurs when the selection changes on a specific table.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-address-member)|Gets the address that represents the changed area of a table on a specific worksheet.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-changetype-member)|Gets the change type that represents how the changed event is triggered.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-source-member)|Gets the source of the event.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-tableid-member)|Gets the ID of the table in which the data changed.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member)|Occurs when data changes on any table in a workbook, or a worksheet.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-address-member)|Gets the range address that represents the selected area of the table on a specific worksheet.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-isinsidetable-member)|Specifies if the selection is inside a table.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-tableid-member)|Gets the ID of the table in which the selection changed.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the selection changed.|
|[Workbook](/javascript/api/excel/excel.workbook)|[dataConnections](/javascript/api/excel/excel.workbook#excel-excel-workbook-dataconnections-member)|Represents all data connections in the workbook.|
||[getActiveCell()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivecell-member(1))|Gets the currently active cell from the workbook.|
||[name](/javascript/api/excel/excel.workbook#excel-excel-workbook-name-member)|Gets the workbook name.|
||[properties](/javascript/api/excel/excel.workbook#excel-excel-workbook-properties-member)|Gets the workbook properties.|
||[protection](/javascript/api/excel/excel.workbook#excel-excel-workbook-protection-member)|Returns the protection object for a workbook.|
||[styles](/javascript/api/excel/excel.workbook#excel-excel-workbook-styles-member)|Represents a collection of styles associated with the workbook.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protect-member(1))|Protects a workbook.|
||[protected](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protected-member)|Specifies if the workbook is protected.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-unprotect-member(1))|Unprotects a workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel.WorksheetPositionType, relativeTo?: Excel.Worksheet)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-copy-member(1))|Copies a worksheet and places it at the specified position.|
||[freezePanes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezepanes-member)|Gets an object that can be used to manipulate frozen panes on the worksheet.|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrangebyindexes-member(1))|Gets the `Range` object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|
||[onActivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member)|Occurs when the worksheet is activated.|
||[onChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member)|Occurs when data changes in a specific worksheet.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)|Occurs when the worksheet is deactivated.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)|Occurs when the selection changes on a specific worksheet.|
||[standardHeight](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardheight-member)|Returns the standard (default) height of all the rows in the worksheet, in points.|
||[standardWidth](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardwidth-member)|Specifies the standard (default) width of all the columns in the worksheet.|
||[tabColor](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabcolor-member)|The tab color of the worksheet.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet that is activated.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-worksheetid-member)|Gets the ID of the worksheet that is added to the workbook.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-address-member)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changetype-member)|Gets the change type that represents how the changed event is triggered.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member)|Occurs when any worksheet in the workbook is activated.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member)|Occurs when a new worksheet is added to the workbook.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member)|Occurs when any worksheet in the workbook is deactivated.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member)|Occurs when a worksheet is deleted from the workbook.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet that is deactivated.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-worksheetid-member)|Gets the ID of the worksheet that is deleted from the workbook.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezeat-member(1))|Sets the frozen cells in the active worksheet view.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezecolumns-member(1))|Freeze the first column or columns of the worksheet in place.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezerows-member(1))|Freeze the top row or rows of the worksheet in place.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocation-member(1))|Gets a range that describes the frozen cells in the active worksheet view.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocationornullobject-member(1))|Gets a range that describes the frozen cells in the active worksheet view.|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-unfreeze-member(1))|Removes all frozen panes in the worksheet.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-unprotect-member(1))|Unprotects a worksheet.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditobjects-member)|Represents the worksheet protection option allowing editing of objects.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditscenarios-member)|Represents the worksheet protection option allowing editing of scenarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-selectionmode-member)|Represents the worksheet protection option of selection mode.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-address-member)|Gets the range address that represents the selected area of a specific worksheet.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the selection changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
