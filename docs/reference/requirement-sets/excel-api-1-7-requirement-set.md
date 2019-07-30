---
title: Excel JavaScript API requirement set 1.7
description: 'Details about the ExcelApi 1.7 requirement set'
ms.date: 07/25/2019
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

The following table lists the APIs in Excel JavaScript API requirement set 1.7. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.7 or earlier, see [Excel APIs in requirement set 1.7 or earlier](/javascript/api/excel?view=excel-js-1.7).

| Class | Fields | Description |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Represents the type of the chart. See Excel.ChartType for details.|
||[id](/javascript/api/excel/excel.chart#id)|The unique id of chart. Read-only.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Represents whether to display all field buttons on a PivotChart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#border)|Represents the border format of chart area, which includes color, linestyle, and weight. Read-only.|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[border](/javascript/api/excel/excel.chartareaformatdata#border)|Represents the border format of chart area, which includes color, linestyle, and weight. Read-only.|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[border](/javascript/api/excel/excel.chartareaformatloadoptions#border)|Represents the border format of chart area, which includes color, linestyle, and weight.|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[border](/javascript/api/excel/excel.chartareaformatupdatedata#border)|Represents the border format of chart area, which includes color, linestyle, and weight.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: "Invalid" \| "Category" \| "Value" \| "Series", group?: "Primary" \| "Secondary")](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Returns the specific axis identified by type and group.|
||[getItem(type: Excel.ChartAxisType, group?: Excel.ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Returns the specific axis identified by type and group.|
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
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[axisGroup](/javascript/api/excel/excel.chartaxisdata#axisgroup)|Represents the group for the specified axis. See Excel.ChartAxisGroup for details. Read-only.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisdata#basetimeunit)|Returns or sets the base unit for the specified category axis.|
||[categoryType](/javascript/api/excel/excel.chartaxisdata#categorytype)|Returns or sets the category axis type.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisdata#customdisplayunit)|Represents the custom axis display unit value. Read-only. To set this property, please use the SetCustomDisplayUnit(double) method.|
||[displayUnit](/javascript/api/excel/excel.chartaxisdata#displayunit)|Represents the axis display unit. See Excel.ChartAxisDisplayUnit for details.|
||[height](/javascript/api/excel/excel.chartaxisdata#height)|Represents the height, in points, of the chart axis. Null if the axis is not visible. Read-only.|
||[left](/javascript/api/excel/excel.chartaxisdata#left)|Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis is not visible. Read-only.|
||[logBase](/javascript/api/excel/excel.chartaxisdata#logbase)|Represents the base of the logarithm when using logarithmic scales.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisdata#majortickmark)|Represents the type of major tick mark for the specified axis. See Excel.ChartAxisTickMark for details.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#majortimeunitscale)|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisdata#minortickmark)|Represents the type of minor tick mark for the specified axis. See Excel.ChartAxisTickMark for details.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#minortimeunitscale)|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisdata#reverseplotorder)|Represents whether Microsoft Excel plots data points from last to first.|
||[scaleType](/javascript/api/excel/excel.chartaxisdata#scaletype)|Represents the value axis scale type. See Excel.ChartAxisScaleType for details.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisdata#showdisplayunitlabel)|Represents whether the axis display unit label is visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisdata#ticklabelposition)|Represents the position of tick-mark labels on the specified axis. See Excel.ChartAxisTickLabelPosition for details.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisdata#ticklabelspacing)|Represents the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisdata#tickmarkspacing)|Represents the number of categories or series between tick marks.|
||[top](/javascript/api/excel/excel.chartaxisdata#top)|Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis is not visible. Read-only.|
||[type](/javascript/api/excel/excel.chartaxisdata#type)|Represents the axis type. See Excel.ChartAxisType for details.|
||[visible](/javascript/api/excel/excel.chartaxisdata#visible)|A boolean value represents the visibility of the axis.|
||[width](/javascript/api/excel/excel.chartaxisdata#width)|Represents the width, in points, of the chart axis. Null if the axis is not visible. Read-only.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[axisGroup](/javascript/api/excel/excel.chartaxisloadoptions#axisgroup)|Represents the group for the specified axis. See Excel.ChartAxisGroup for details. Read-only.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisloadoptions#basetimeunit)|Returns or sets the base unit for the specified category axis.|
||[categoryType](/javascript/api/excel/excel.chartaxisloadoptions#categorytype)|Returns or sets the category axis type.|
||[crosses](/javascript/api/excel/excel.chartaxisloadoptions#crosses)|[DEPRECATED; kept for back-compat with existing first-party solutions]. Please use `Position` instead.|
||[crossesAt](/javascript/api/excel/excel.chartaxisloadoptions#crossesat)|[DEPRECATED; kept for back-compat with existing first-party solutions]. Please use `PositionAt` instead.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisloadoptions#customdisplayunit)|Represents the custom axis display unit value. Read-only. To set this property, please use the SetCustomDisplayUnit(double) method.|
||[displayUnit](/javascript/api/excel/excel.chartaxisloadoptions#displayunit)|Represents the axis display unit. See Excel.ChartAxisDisplayUnit for details.|
||[height](/javascript/api/excel/excel.chartaxisloadoptions#height)|Represents the height, in points, of the chart axis. Null if the axis is not visible. Read-only.|
||[left](/javascript/api/excel/excel.chartaxisloadoptions#left)|Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis is not visible. Read-only.|
||[logBase](/javascript/api/excel/excel.chartaxisloadoptions#logbase)|Represents the base of the logarithm when using logarithmic scales.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#majortickmark)|Represents the type of major tick mark for the specified axis. See Excel.ChartAxisTickMark for details.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#majortimeunitscale)|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#minortickmark)|Represents the type of minor tick mark for the specified axis. See Excel.ChartAxisTickMark for details.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#minortimeunitscale)|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisloadoptions#reverseplotorder)|Represents whether Microsoft Excel plots data points from last to first.|
||[scaleType](/javascript/api/excel/excel.chartaxisloadoptions#scaletype)|Represents the value axis scale type. See Excel.ChartAxisScaleType for details.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisloadoptions#showdisplayunitlabel)|Represents whether the axis display unit label is visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelposition)|Represents the position of tick-mark labels on the specified axis. See Excel.ChartAxisTickLabelPosition for details.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelspacing)|Represents the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisloadoptions#tickmarkspacing)|Represents the number of categories or series between tick marks.|
||[top](/javascript/api/excel/excel.chartaxisloadoptions#top)|Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis is not visible. Read-only.|
||[type](/javascript/api/excel/excel.chartaxisloadoptions#type)|Represents the axis type. See Excel.ChartAxisType for details.|
||[visible](/javascript/api/excel/excel.chartaxisloadoptions#visible)|A boolean value represents the visibility of the axis.|
||[width](/javascript/api/excel/excel.chartaxisloadoptions#width)|Represents the width, in points, of the chart axis. Null if the axis is not visible. Read-only.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[baseTimeUnit](/javascript/api/excel/excel.chartaxisupdatedata#basetimeunit)|Returns or sets the base unit for the specified category axis.|
||[categoryType](/javascript/api/excel/excel.chartaxisupdatedata#categorytype)|Returns or sets the category axis type.|
||[displayUnit](/javascript/api/excel/excel.chartaxisupdatedata#displayunit)|Represents the axis display unit. See Excel.ChartAxisDisplayUnit for details.|
||[logBase](/javascript/api/excel/excel.chartaxisupdatedata#logbase)|Represents the base of the logarithm when using logarithmic scales.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#majortickmark)|Represents the type of major tick mark for the specified axis. See Excel.ChartAxisTickMark for details.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#majortimeunitscale)|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#minortickmark)|Represents the type of minor tick mark for the specified axis. See Excel.ChartAxisTickMark for details.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#minortimeunitscale)|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisupdatedata#reverseplotorder)|Represents whether Microsoft Excel plots data points from last to first.|
||[scaleType](/javascript/api/excel/excel.chartaxisupdatedata#scaletype)|Represents the value axis scale type. See Excel.ChartAxisScaleType for details.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisupdatedata#showdisplayunitlabel)|Represents whether the axis display unit label is visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelposition)|Represents the position of tick-mark labels on the specified axis. See Excel.ChartAxisTickLabelPosition for details.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelspacing)|Represents the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisupdatedata#tickmarkspacing)|Represents the number of categories or series between tick marks.|
||[visible](/javascript/api/excel/excel.chartaxisupdatedata#visible)|A boolean value represents the visibility of the axis.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|HTML color code representing the color of borders in the chart.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Represents the line style of the border. See Excel.ChartLineStyle for details.|
||[set(properties: Excel.ChartBorder)](/javascript/api/excel/excel.chartborder#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartBorderUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartborder#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Represents weight of the border, in points.|
|[ChartBorderData](/javascript/api/excel/excel.chartborderdata)|[color](/javascript/api/excel/excel.chartborderdata#color)|HTML color code representing the color of borders in the chart.|
||[lineStyle](/javascript/api/excel/excel.chartborderdata#linestyle)|Represents the line style of the border. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartborderdata#weight)|Represents weight of the border, in points.|
|[ChartBorderLoadOptions](/javascript/api/excel/excel.chartborderloadoptions)|[$all](/javascript/api/excel/excel.chartborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartborderloadoptions#color)|HTML color code representing the color of borders in the chart.|
||[lineStyle](/javascript/api/excel/excel.chartborderloadoptions#linestyle)|Represents the line style of the border. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartborderloadoptions#weight)|Represents weight of the border, in points.|
|[ChartBorderUpdateData](/javascript/api/excel/excel.chartborderupdatedata)|[color](/javascript/api/excel/excel.chartborderupdatedata#color)|HTML color code representing the color of borders in the chart.|
||[lineStyle](/javascript/api/excel/excel.chartborderupdatedata#linestyle)|Represents the line style of the border. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartborderupdatedata#weight)|Represents weight of the border, in points.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartcollectionloadoptions#charttype)|For EACH ITEM in the collection: Represents the type of the chart. See Excel.ChartType for details.|
||[id](/javascript/api/excel/excel.chartcollectionloadoptions#id)|For EACH ITEM in the collection: The unique id of chart. Read-only.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartcollectionloadoptions#showallfieldbuttons)|For EACH ITEM in the collection: Represents whether to display all field buttons on a PivotChart.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[chartType](/javascript/api/excel/excel.chartdata#charttype)|Represents the type of the chart. See Excel.ChartType for details.|
||[id](/javascript/api/excel/excel.chartdata#id)|The unique id of chart. Read-only.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartdata#showallfieldbuttons)|Represents whether to display all field buttons on a PivotChart.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[separator](/javascript/api/excel/excel.chartdatalabel#separator)|String representing the separator used for the data label on a chart.|
||[set(properties: Excel.ChartDataLabel)](/javascript/api/excel/excel.chartdatalabel#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartDataLabelUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartdatalabel#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[position](/javascript/api/excel/excel.chartdatalabeldata#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[separator](/javascript/api/excel/excel.chartdatalabeldata#separator)|String representing the separator used for the data label on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabeldata#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabeldata#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabeldata#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabeldata#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabeldata#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabeldata#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelloadoptions#$all)||
||[position](/javascript/api/excel/excel.chartdatalabelloadoptions#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[separator](/javascript/api/excel/excel.chartdatalabelloadoptions#separator)|String representing the separator used for the data label on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelloadoptions#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelloadoptions#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelloadoptions#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelloadoptions#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelloadoptions#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabelloadoptions#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[position](/javascript/api/excel/excel.chartdatalabelupdatedata#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[separator](/javascript/api/excel/excel.chartdatalabelupdatedata#separator)|String representing the separator used for the data label on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelupdatedata#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelupdatedata#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelupdatedata#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelupdatedata#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelupdatedata#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabelupdatedata#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Represents the font attributes, such as font name, font size, color, etc. of chart characters object.|
||[set(properties: Excel.ChartFormatString)](/javascript/api/excel/excel.chartformatstring#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartFormatStringUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartformatstring#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartFormatStringData](/javascript/api/excel/excel.chartformatstringdata)|[font](/javascript/api/excel/excel.chartformatstringdata#font)|Represents the font attributes, such as font name, font size, color, etc. of chart characters object.|
|[ChartFormatStringLoadOptions](/javascript/api/excel/excel.chartformatstringloadoptions)|[$all](/javascript/api/excel/excel.chartformatstringloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartformatstringloadoptions#font)|Represents the font attributes, such as font name, font size, color, etc. of chart characters object.|
|[ChartFormatStringUpdateData](/javascript/api/excel/excel.chartformatstringupdatedata)|[font](/javascript/api/excel/excel.chartformatstringupdatedata#font)|Represents the font attributes, such as font name, font size, color, etc. of chart characters object.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Represents the height, in points, of the legend on the chart. Null if legend is not visible.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Represents the left, in points, of a chart legend. Null if legend is not visible.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Represents a collection of legendEntries in the legend. Read-only.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|Represents if the legend has a shadow on the chart.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Represents the top of a chart legend.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Represents the width, in points, of the legend on the chart. Null if legend is not visible.|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[height](/javascript/api/excel/excel.chartlegenddata#height)|Represents the height, in points, of the legend on the chart. Null if legend is not visible.|
||[left](/javascript/api/excel/excel.chartlegenddata#left)|Represents the left, in points, of a chart legend. Null if legend is not visible.|
||[legendEntries](/javascript/api/excel/excel.chartlegenddata#legendentries)|Represents a collection of legendEntries in the legend. Read-only.|
||[showShadow](/javascript/api/excel/excel.chartlegenddata#showshadow)|Represents if the legend has a shadow on the chart.|
||[top](/javascript/api/excel/excel.chartlegenddata#top)|Represents the top of a chart legend.|
||[width](/javascript/api/excel/excel.chartlegenddata#width)|Represents the width, in points, of the legend on the chart. Null if legend is not visible.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[set(properties: Excel.ChartLegendEntry)](/javascript/api/excel/excel.chartlegendentry#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartLegendEntryUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartlegendentry#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Represents the visible of a chart legend entry.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Returns the number of legendEntry in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Returns a legendEntry at the given index.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Gets the loaded child items in this collection.|
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#visible)|For EACH ITEM in the collection: Represents the visible of a chart legend entry.|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[visible](/javascript/api/excel/excel.chartlegendentrydata#visible)|Represents the visible of a chart legend entry.|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentryloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentryloadoptions#visible)|Represents the visible of a chart legend entry.|
|[ChartLegendEntryUpdateData](/javascript/api/excel/excel.chartlegendentryupdatedata)|[visible](/javascript/api/excel/excel.chartlegendentryupdatedata#visible)|Represents the visible of a chart legend entry.|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[height](/javascript/api/excel/excel.chartlegendloadoptions#height)|Represents the height, in points, of the legend on the chart. Null if legend is not visible.|
||[left](/javascript/api/excel/excel.chartlegendloadoptions#left)|Represents the left, in points, of a chart legend. Null if legend is not visible.|
||[showShadow](/javascript/api/excel/excel.chartlegendloadoptions#showshadow)|Represents if the legend has a shadow on the chart.|
||[top](/javascript/api/excel/excel.chartlegendloadoptions#top)|Represents the top of a chart legend.|
||[width](/javascript/api/excel/excel.chartlegendloadoptions#width)|Represents the width, in points, of the legend on the chart. Null if legend is not visible.|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[height](/javascript/api/excel/excel.chartlegendupdatedata#height)|Represents the height, in points, of the legend on the chart. Null if legend is not visible.|
||[left](/javascript/api/excel/excel.chartlegendupdatedata#left)|Represents the left, in points, of a chart legend. Null if legend is not visible.|
||[showShadow](/javascript/api/excel/excel.chartlegendupdatedata#showshadow)|Represents if the legend has a shadow on the chart.|
||[top](/javascript/api/excel/excel.chartlegendupdatedata#top)|Represents the top of a chart legend.|
||[width](/javascript/api/excel/excel.chartlegendupdatedata#width)|Represents the width, in points, of the legend on the chart. Null if legend is not visible.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Represents the line style. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Represents weight of the line, in points.|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[lineStyle](/javascript/api/excel/excel.chartlineformatdata#linestyle)|Represents the line style. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartlineformatdata#weight)|Represents weight of the line, in points.|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[lineStyle](/javascript/api/excel/excel.chartlineformatloadoptions#linestyle)|Represents the line style. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartlineformatloadoptions#weight)|Represents weight of the line, in points.|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[lineStyle](/javascript/api/excel/excel.chartlineformatupdatedata#linestyle)|Represents the line style. See Excel.ChartLineStyle for details.|
||[weight](/javascript/api/excel/excel.chartlineformatupdatedata#weight)|Represents weight of the line, in points.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[chartType](/javascript/api/excel/excel.chartloadoptions#charttype)|Represents the type of the chart. See Excel.ChartType for details.|
||[id](/javascript/api/excel/excel.chartloadoptions#id)|The unique id of chart. Read-only.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartloadoptions#showallfieldbuttons)|Represents whether to display all field buttons on a PivotChart.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Represents whether a data point has a data label. Not applicable for surface charts.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Represents marker size of data point.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Represents marker style of a chart data point. See Excel.ChartMarkerStyle for details.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Returns the data label of a chart point. Read-only.|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[dataLabel](/javascript/api/excel/excel.chartpointdata#datalabel)|Returns the data label of a chart point. Read-only.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointdata#hasdatalabel)|Represents whether a data point has a data label. Not applicable for surface charts.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointdata#markerbackgroundcolor)|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointdata#markerforegroundcolor)|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|
||[markerSize](/javascript/api/excel/excel.chartpointdata#markersize)|Represents marker size of data point.|
||[markerStyle](/javascript/api/excel/excel.chartpointdata#markerstyle)|Represents marker style of a chart data point. See Excel.ChartMarkerStyle for details.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#border)|Represents the border format of a chart data point, which includes color, style, and weight information. Read-only.|
|[ChartPointFormatData](/javascript/api/excel/excel.chartpointformatdata)|[border](/javascript/api/excel/excel.chartpointformatdata#border)|Represents the border format of a chart data point, which includes color, style, and weight information. Read-only.|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[border](/javascript/api/excel/excel.chartpointformatloadoptions#border)|Represents the border format of a chart data point, which includes color, style, and weight information.|
|[ChartPointFormatUpdateData](/javascript/api/excel/excel.chartpointformatupdatedata)|[border](/javascript/api/excel/excel.chartpointformatupdatedata#border)|Represents the border format of a chart data point, which includes color, style, and weight information.|
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointloadoptions#datalabel)|Returns the data label of a chart point.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointloadoptions#hasdatalabel)|Represents whether a data point has a data label. Not applicable for surface charts.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerbackgroundcolor)|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerforegroundcolor)|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|
||[markerSize](/javascript/api/excel/excel.chartpointloadoptions#markersize)|Represents marker size of data point.|
||[markerStyle](/javascript/api/excel/excel.chartpointloadoptions#markerstyle)|Represents marker style of a chart data point. See Excel.ChartMarkerStyle for details.|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[dataLabel](/javascript/api/excel/excel.chartpointupdatedata#datalabel)|Returns the data label of a chart point.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointupdatedata#hasdatalabel)|Represents whether a data point has a data label. Not applicable for surface charts.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerbackgroundcolor)|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerforegroundcolor)|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|
||[markerSize](/javascript/api/excel/excel.chartpointupdatedata#markersize)|Represents marker size of data point.|
||[markerStyle](/javascript/api/excel/excel.chartpointupdatedata#markerstyle)|Represents marker style of a chart data point. See Excel.ChartMarkerStyle for details.|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#datalabel)|For EACH ITEM in the collection: Returns the data label of a chart point.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#hasdatalabel)|For EACH ITEM in the collection: Represents whether a data point has a data label. Not applicable for surface charts.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerbackgroundcolor)|For EACH ITEM in the collection: HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerforegroundcolor)|For EACH ITEM in the collection: HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|
||[markerSize](/javascript/api/excel/excel.chartpointscollectionloadoptions#markersize)|For EACH ITEM in the collection: Represents marker size of data point.|
||[markerStyle](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerstyle)|For EACH ITEM in the collection: Represents marker style of a chart data point. See Excel.ChartMarkerStyle for details.|
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
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartseriescollectionloadoptions#charttype)|For EACH ITEM in the collection: Represents the chart type of a series. See Excel.ChartType for details.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#doughnutholesize)|For EACH ITEM in the collection: Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|
||[filtered](/javascript/api/excel/excel.chartseriescollectionloadoptions#filtered)|For EACH ITEM in the collection: Boolean value representing if the series is filtered or not. Not applicable for surface charts.|
||[gapWidth](/javascript/api/excel/excel.chartseriescollectionloadoptions#gapwidth)|For EACH ITEM in the collection: Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|
||[hasDataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#hasdatalabels)|For EACH ITEM in the collection: Boolean value representing if the series has data labels or not.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerbackgroundcolor)|For EACH ITEM in the collection: Represents markers background color of a chart series.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerforegroundcolor)|For EACH ITEM in the collection: Represents markers foreground color of a chart series.|
||[markerSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#markersize)|For EACH ITEM in the collection: Represents marker size of a chart series.|
||[markerStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerstyle)|For EACH ITEM in the collection: Represents marker style of a chart series. See Excel.ChartMarkerStyle for details.|
||[plotOrder](/javascript/api/excel/excel.chartseriescollectionloadoptions#plotorder)|For EACH ITEM in the collection: Represents the plot order of a chart series within the chart group.|
||[showShadow](/javascript/api/excel/excel.chartseriescollectionloadoptions#showshadow)|For EACH ITEM in the collection: Boolean value representing if the series has a shadow or not.|
||[smooth](/javascript/api/excel/excel.chartseriescollectionloadoptions#smooth)|For EACH ITEM in the collection: Boolean value representing if the series is smooth or not. Only applicable to line and scatter charts.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[chartType](/javascript/api/excel/excel.chartseriesdata#charttype)|Represents the chart type of a series. See Excel.ChartType for details.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesdata#doughnutholesize)|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|
||[filtered](/javascript/api/excel/excel.chartseriesdata#filtered)|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|
||[gapWidth](/javascript/api/excel/excel.chartseriesdata#gapwidth)|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesdata#hasdatalabels)|Boolean value representing if the series has data labels or not.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesdata#markerbackgroundcolor)|Represents markers background color of a chart series.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesdata#markerforegroundcolor)|Represents markers foreground color of a chart series.|
||[markerSize](/javascript/api/excel/excel.chartseriesdata#markersize)|Represents marker size of a chart series.|
||[markerStyle](/javascript/api/excel/excel.chartseriesdata#markerstyle)|Represents marker style of a chart series. See Excel.ChartMarkerStyle for details.|
||[plotOrder](/javascript/api/excel/excel.chartseriesdata#plotorder)|Represents the plot order of a chart series within the chart group.|
||[showShadow](/javascript/api/excel/excel.chartseriesdata#showshadow)|Boolean value representing if the series has a shadow or not.|
||[smooth](/javascript/api/excel/excel.chartseriesdata#smooth)|Boolean value representing if the series is smooth or not. Only applicable to line and scatter charts.|
||[trendlines](/javascript/api/excel/excel.chartseriesdata#trendlines)|Represents a collection of trendlines in the series. Read-only.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[chartType](/javascript/api/excel/excel.chartseriesloadoptions#charttype)|Represents the chart type of a series. See Excel.ChartType for details.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesloadoptions#doughnutholesize)|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|
||[filtered](/javascript/api/excel/excel.chartseriesloadoptions#filtered)|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|
||[gapWidth](/javascript/api/excel/excel.chartseriesloadoptions#gapwidth)|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesloadoptions#hasdatalabels)|Boolean value representing if the series has data labels or not.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerbackgroundcolor)|Represents markers background color of a chart series.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerforegroundcolor)|Represents markers foreground color of a chart series.|
||[markerSize](/javascript/api/excel/excel.chartseriesloadoptions#markersize)|Represents marker size of a chart series.|
||[markerStyle](/javascript/api/excel/excel.chartseriesloadoptions#markerstyle)|Represents marker style of a chart series. See Excel.ChartMarkerStyle for details.|
||[plotOrder](/javascript/api/excel/excel.chartseriesloadoptions#plotorder)|Represents the plot order of a chart series within the chart group.|
||[showShadow](/javascript/api/excel/excel.chartseriesloadoptions#showshadow)|Boolean value representing if the series has a shadow or not.|
||[smooth](/javascript/api/excel/excel.chartseriesloadoptions#smooth)|Boolean value representing if the series is smooth or not. Only applicable to line and scatter charts.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[chartType](/javascript/api/excel/excel.chartseriesupdatedata#charttype)|Represents the chart type of a series. See Excel.ChartType for details.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesupdatedata#doughnutholesize)|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|
||[filtered](/javascript/api/excel/excel.chartseriesupdatedata#filtered)|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|
||[gapWidth](/javascript/api/excel/excel.chartseriesupdatedata#gapwidth)|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesupdatedata#hasdatalabels)|Boolean value representing if the series has data labels or not.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerbackgroundcolor)|Represents markers background color of a chart series.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerforegroundcolor)|Represents markers foreground color of a chart series.|
||[markerSize](/javascript/api/excel/excel.chartseriesupdatedata#markersize)|Represents marker size of a chart series.|
||[markerStyle](/javascript/api/excel/excel.chartseriesupdatedata#markerstyle)|Represents marker style of a chart series. See Excel.ChartMarkerStyle for details.|
||[plotOrder](/javascript/api/excel/excel.chartseriesupdatedata#plotorder)|Represents the plot order of a chart series within the chart group.|
||[showShadow](/javascript/api/excel/excel.chartseriesupdatedata#showshadow)|Boolean value representing if the series has a shadow or not.|
||[smooth](/javascript/api/excel/excel.chartseriesupdatedata#smooth)|Boolean value representing if the series is smooth or not. Only applicable to line and scatter charts.|
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
|[ChartTitleData](/javascript/api/excel/excel.charttitledata)|[height](/javascript/api/excel/excel.charttitledata#height)|Returns the height, in points, of the chart title. Null if chart title is not visible. Read-only.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitledata#horizontalalignment)|Represents the horizontal alignment for chart title.|
||[left](/javascript/api/excel/excel.charttitledata#left)|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title is not visible.|
||[position](/javascript/api/excel/excel.charttitledata#position)|Represents the position of chart title. See Excel.ChartTitlePosition for details.|
||[showShadow](/javascript/api/excel/excel.charttitledata#showshadow)|Represents a boolean value that determines if the chart title has a shadow.|
||[textOrientation](/javascript/api/excel/excel.charttitledata#textorientation)|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttitledata#top)|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitledata#verticalalignment)|Represents the vertical alignment of chart title. See Excel.ChartTextVerticalAlignment for details.|
||[width](/javascript/api/excel/excel.charttitledata#width)|Returns the width, in points, of the chart title. Null if chart title is not visible. Read-only.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#border)|Represents the border format of chart title, which includes color, linestyle, and weight. Read-only.|
|[ChartTitleFormatData](/javascript/api/excel/excel.charttitleformatdata)|[border](/javascript/api/excel/excel.charttitleformatdata#border)|Represents the border format of chart title, which includes color, linestyle, and weight. Read-only.|
|[ChartTitleFormatLoadOptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[border](/javascript/api/excel/excel.charttitleformatloadoptions#border)|Represents the border format of chart title, which includes color, linestyle, and weight.|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[border](/javascript/api/excel/excel.charttitleformatupdatedata#border)|Represents the border format of chart title, which includes color, linestyle, and weight.|
|[ChartTitleLoadOptions](/javascript/api/excel/excel.charttitleloadoptions)|[height](/javascript/api/excel/excel.charttitleloadoptions#height)|Returns the height, in points, of the chart title. Null if chart title is not visible. Read-only.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitleloadoptions#horizontalalignment)|Represents the horizontal alignment for chart title.|
||[left](/javascript/api/excel/excel.charttitleloadoptions#left)|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title is not visible.|
||[position](/javascript/api/excel/excel.charttitleloadoptions#position)|Represents the position of chart title. See Excel.ChartTitlePosition for details.|
||[showShadow](/javascript/api/excel/excel.charttitleloadoptions#showshadow)|Represents a boolean value that determines if the chart title has a shadow.|
||[textOrientation](/javascript/api/excel/excel.charttitleloadoptions#textorientation)|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttitleloadoptions#top)|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitleloadoptions#verticalalignment)|Represents the vertical alignment of chart title. See Excel.ChartTextVerticalAlignment for details.|
||[width](/javascript/api/excel/excel.charttitleloadoptions#width)|Returns the width, in points, of the chart title. Null if chart title is not visible. Read-only.|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[horizontalAlignment](/javascript/api/excel/excel.charttitleupdatedata#horizontalalignment)|Represents the horizontal alignment for chart title.|
||[left](/javascript/api/excel/excel.charttitleupdatedata#left)|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title is not visible.|
||[position](/javascript/api/excel/excel.charttitleupdatedata#position)|Represents the position of chart title. See Excel.ChartTitlePosition for details.|
||[showShadow](/javascript/api/excel/excel.charttitleupdatedata#showshadow)|Represents a boolean value that determines if the chart title has a shadow.|
||[textOrientation](/javascript/api/excel/excel.charttitleupdatedata#textorientation)|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttitleupdatedata#top)|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitleupdatedata#verticalalignment)|Represents the vertical alignment of chart title. See Excel.ChartTextVerticalAlignment for details.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Delete the trendline object.|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Represents the formatting of a chart trendline.|
||[set(properties: Excel.ChartTrendline)](/javascript/api/excel/excel.charttrendline#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartTrendlineUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.charttrendline#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Represents the type of a chart trendline.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: "Linear" \| "Exponential" \| "Logarithmic" \| "MovingAverage" \| "Polynomial" \| "Power")](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Adds a new trendline to trendline collection.|
||[add(type?: Excel.ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Adds a new trendline to trendline collection.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Returns the number of trendlines in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Get trendline object by index, which is the insertion order in items array.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Gets the loaded child items in this collection.|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#format)|For EACH ITEM in the collection: Represents the formatting of a chart trendline.|
||[intercept](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#intercept)|For EACH ITEM in the collection: Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#movingaverageperiod)|For EACH ITEM in the collection: Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.|
||[name](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#name)|For EACH ITEM in the collection: Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#polynomialorder)|For EACH ITEM in the collection: Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.|
||[type](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#type)|For EACH ITEM in the collection: Represents the type of a chart trendline.|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[format](/javascript/api/excel/excel.charttrendlinedata#format)|Represents the formatting of a chart trendline.|
||[intercept](/javascript/api/excel/excel.charttrendlinedata#intercept)|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinedata#movingaverageperiod)|Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.|
||[name](/javascript/api/excel/excel.charttrendlinedata#name)|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinedata#polynomialorder)|Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.|
||[type](/javascript/api/excel/excel.charttrendlinedata#type)|Represents the type of a chart trendline.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Represents chart line formatting. Read-only.|
||[set(properties: Excel.ChartTrendlineFormat)](/javascript/api/excel/excel.charttrendlineformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartTrendlineFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.charttrendlineformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartTrendlineFormatData](/javascript/api/excel/excel.charttrendlineformatdata)|[line](/javascript/api/excel/excel.charttrendlineformatdata#line)|Represents chart line formatting. Read-only.|
|[ChartTrendlineFormatLoadOptions](/javascript/api/excel/excel.charttrendlineformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charttrendlineformatloadoptions#line)|Represents chart line formatting.|
|[ChartTrendlineFormatUpdateData](/javascript/api/excel/excel.charttrendlineformatupdatedata)|[line](/javascript/api/excel/excel.charttrendlineformatupdatedata#line)|Represents chart line formatting.|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlineloadoptions#format)|Represents the formatting of a chart trendline.|
||[intercept](/javascript/api/excel/excel.charttrendlineloadoptions#intercept)|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineloadoptions#movingaverageperiod)|Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.|
||[name](/javascript/api/excel/excel.charttrendlineloadoptions#name)|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineloadoptions#polynomialorder)|Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.|
||[type](/javascript/api/excel/excel.charttrendlineloadoptions#type)|Represents the type of a chart trendline.|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[format](/javascript/api/excel/excel.charttrendlineupdatedata#format)|Represents the formatting of a chart trendline.|
||[intercept](/javascript/api/excel/excel.charttrendlineupdatedata#intercept)|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineupdatedata#movingaverageperiod)|Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.|
||[name](/javascript/api/excel/excel.charttrendlineupdatedata#name)|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineupdatedata#polynomialorder)|Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.|
||[type](/javascript/api/excel/excel.charttrendlineupdatedata#type)|Represents the type of a chart trendline.|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[chartType](/javascript/api/excel/excel.chartupdatedata#charttype)|Represents the type of the chart. See Excel.ChartType for details.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartupdatedata#showallfieldbuttons)|Represents whether to display all field buttons on a PivotChart.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Deletes the custom property.|
||[key](/javascript/api/excel/excel.customproperty#key)|Gets the key of the custom property. Read only.|
||[type](/javascript/api/excel/excel.customproperty#type)|Gets the value type of the custom property. Read only.|
||[set(properties: Excel.CustomProperty)](/javascript/api/excel/excel.customproperty#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.CustomPropertyUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.customproperty#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[value](/javascript/api/excel/excel.customproperty#value)|Gets or sets the value of the custom property.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Gets the loaded child items in this collection.|
|[CustomPropertyCollectionLoadOptions](/javascript/api/excel/excel.custompropertycollectionloadoptions)|[$all](/javascript/api/excel/excel.custompropertycollectionloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertycollectionloadoptions#key)|For EACH ITEM in the collection: Gets the key of the custom property. Read only.|
||[type](/javascript/api/excel/excel.custompropertycollectionloadoptions#type)|For EACH ITEM in the collection: Gets the value type of the custom property. Read only.|
||[value](/javascript/api/excel/excel.custompropertycollectionloadoptions#value)|For EACH ITEM in the collection: Gets or sets the value of the custom property.|
|[CustomPropertyData](/javascript/api/excel/excel.custompropertydata)|[key](/javascript/api/excel/excel.custompropertydata#key)|Gets the key of the custom property. Read only.|
||[type](/javascript/api/excel/excel.custompropertydata#type)|Gets the value type of the custom property. Read only.|
||[value](/javascript/api/excel/excel.custompropertydata#value)|Gets or sets the value of the custom property.|
|[CustomPropertyLoadOptions](/javascript/api/excel/excel.custompropertyloadoptions)|[$all](/javascript/api/excel/excel.custompropertyloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertyloadoptions#key)|Gets the key of the custom property. Read only.|
||[type](/javascript/api/excel/excel.custompropertyloadoptions#type)|Gets the value type of the custom property. Read only.|
||[value](/javascript/api/excel/excel.custompropertyloadoptions#value)|Gets or sets the value of the custom property.|
|[CustomPropertyUpdateData](/javascript/api/excel/excel.custompropertyupdatedata)|[value](/javascript/api/excel/excel.custompropertyupdatedata#value)|Gets or sets the value of the custom property.|
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
||[set(properties: Excel.DocumentProperties)](/javascript/api/excel/excel.documentproperties#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.DocumentPropertiesUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.documentproperties#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Gets or sets the subject of the workbook.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Gets or sets the title of the workbook.|
|[DocumentPropertiesData](/javascript/api/excel/excel.documentpropertiesdata)|[author](/javascript/api/excel/excel.documentpropertiesdata#author)|Gets or sets the author of the workbook.|
||[category](/javascript/api/excel/excel.documentpropertiesdata#category)|Gets or sets the category of the workbook.|
||[comments](/javascript/api/excel/excel.documentpropertiesdata#comments)|Gets or sets the comments of the workbook.|
||[company](/javascript/api/excel/excel.documentpropertiesdata#company)|Gets or sets the company of the workbook.|
||[creationDate](/javascript/api/excel/excel.documentpropertiesdata#creationdate)|Gets the creation date of the workbook. Read only.|
||[custom](/javascript/api/excel/excel.documentpropertiesdata#custom)|Gets the collection of custom properties of the workbook. Read only.|
||[keywords](/javascript/api/excel/excel.documentpropertiesdata#keywords)|Gets or sets the keywords of the workbook.|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesdata#lastauthor)|Gets the last author of the workbook. Read only.|
||[manager](/javascript/api/excel/excel.documentpropertiesdata#manager)|Gets or sets the manager of the workbook.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesdata#revisionnumber)|Gets the revision number of the workbook. Read only.|
||[subject](/javascript/api/excel/excel.documentpropertiesdata#subject)|Gets or sets the subject of the workbook.|
||[title](/javascript/api/excel/excel.documentpropertiesdata#title)|Gets or sets the title of the workbook.|
|[DocumentPropertiesLoadOptions](/javascript/api/excel/excel.documentpropertiesloadoptions)|[$all](/javascript/api/excel/excel.documentpropertiesloadoptions#$all)||
||[author](/javascript/api/excel/excel.documentpropertiesloadoptions#author)|Gets or sets the author of the workbook.|
||[category](/javascript/api/excel/excel.documentpropertiesloadoptions#category)|Gets or sets the category of the workbook.|
||[comments](/javascript/api/excel/excel.documentpropertiesloadoptions#comments)|Gets or sets the comments of the workbook.|
||[company](/javascript/api/excel/excel.documentpropertiesloadoptions#company)|Gets or sets the company of the workbook.|
||[creationDate](/javascript/api/excel/excel.documentpropertiesloadoptions#creationdate)|Gets the creation date of the workbook. Read only.|
||[keywords](/javascript/api/excel/excel.documentpropertiesloadoptions#keywords)|Gets or sets the keywords of the workbook.|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesloadoptions#lastauthor)|Gets the last author of the workbook. Read only.|
||[manager](/javascript/api/excel/excel.documentpropertiesloadoptions#manager)|Gets or sets the manager of the workbook.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesloadoptions#revisionnumber)|Gets the revision number of the workbook. Read only.|
||[subject](/javascript/api/excel/excel.documentpropertiesloadoptions#subject)|Gets or sets the subject of the workbook.|
||[title](/javascript/api/excel/excel.documentpropertiesloadoptions#title)|Gets or sets the title of the workbook.|
|[DocumentPropertiesUpdateData](/javascript/api/excel/excel.documentpropertiesupdatedata)|[author](/javascript/api/excel/excel.documentpropertiesupdatedata#author)|Gets or sets the author of the workbook.|
||[category](/javascript/api/excel/excel.documentpropertiesupdatedata#category)|Gets or sets the category of the workbook.|
||[comments](/javascript/api/excel/excel.documentpropertiesupdatedata#comments)|Gets or sets the comments of the workbook.|
||[company](/javascript/api/excel/excel.documentpropertiesupdatedata#company)|Gets or sets the company of the workbook.|
||[keywords](/javascript/api/excel/excel.documentpropertiesupdatedata#keywords)|Gets or sets the keywords of the workbook.|
||[manager](/javascript/api/excel/excel.documentpropertiesupdatedata#manager)|Gets or sets the manager of the workbook.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesupdatedata#revisionnumber)|Gets the revision number of the workbook. Read only.|
||[subject](/javascript/api/excel/excel.documentpropertiesupdatedata#subject)|Gets or sets the subject of the workbook.|
||[title](/javascript/api/excel/excel.documentpropertiesupdatedata#title)|Gets or sets the title of the workbook.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Returns an object containing values and types of the named item. Read-only.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Represents the types for each item in the named item array|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Represents the values of each item in the named item array.|
|[NamedItemArrayValuesData](/javascript/api/excel/excel.nameditemarrayvaluesdata)|[types](/javascript/api/excel/excel.nameditemarrayvaluesdata#types)|Represents the types for each item in the named item array|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesdata#values)|Represents the values of each item in the named item array.|
|[NamedItemArrayValuesLoadOptions](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions)|[$all](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#$all)||
||[types](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#types)|Represents the types for each item in the named item array|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#values)|Represents the values of each item in the named item array.|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemcollectionloadoptions#arrayvalues)|For EACH ITEM in the collection: Returns an object containing values and types of the named item.|
||[formula](/javascript/api/excel/excel.nameditemcollectionloadoptions#formula)|For EACH ITEM in the collection: Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[arrayValues](/javascript/api/excel/excel.nameditemdata#arrayvalues)|Returns an object containing values and types of the named item. Read-only.|
||[formula](/javascript/api/excel/excel.nameditemdata#formula)|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemloadoptions#arrayvalues)|Returns an object containing values and types of the named item.|
||[formula](/javascript/api/excel/excel.nameditemloadoptions#formula)|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[formula](/javascript/api/excel/excel.nameditemupdatedata#formula)|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|
||[getImage()](/javascript/api/excel/excel.range#getimage--)|Renders the range as a base64-encoded png image.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Represents the hyperlink for the current range.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Represents Excel's number format code for the given range as a string in the language of the user.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Represents if the current range is an entire column. Read-only.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Represents if the current range is an entire row. Read-only.|
||[showCard()](/javascript/api/excel/excel.range#showcard--)|Displays the card for an active cell if it has rich value content.|
||[style](/javascript/api/excel/excel.range#style)|Represents the style of the current range.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hyperlink](/javascript/api/excel/excel.rangedata#hyperlink)|Represents the hyperlink for the current range.|
||[isEntireColumn](/javascript/api/excel/excel.rangedata#isentirecolumn)|Represents if the current range is an entire column. Read-only.|
||[isEntireRow](/javascript/api/excel/excel.rangedata#isentirerow)|Represents if the current range is an entire row. Read-only.|
||[numberFormatLocal](/javascript/api/excel/excel.rangedata#numberformatlocal)|Represents Excel's number format code for the given range as a string in the language of the user.|
||[style](/javascript/api/excel/excel.rangedata#style)|Represents the style of the current range.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Gets or sets the text orientation of all the cells within the range.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Determines if the row height of the Range object equals the standard height of the sheet.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Indicates whether the column width of the Range object equals the standard width of the sheet.|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[textOrientation](/javascript/api/excel/excel.rangeformatdata#textorientation)|Gets or sets the text orientation of all the cells within the range.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatdata#usestandardheight)|Determines if the row height of the Range object equals the standard height of the sheet.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatdata#usestandardwidth)|Indicates whether the column width of the Range object equals the standard width of the sheet.|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[textOrientation](/javascript/api/excel/excel.rangeformatloadoptions#textorientation)|Gets or sets the text orientation of all the cells within the range.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatloadoptions#usestandardheight)|Determines if the row height of the Range object equals the standard height of the sheet.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatloadoptions#usestandardwidth)|Indicates whether the column width of the Range object equals the standard width of the sheet.|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[textOrientation](/javascript/api/excel/excel.rangeformatupdatedata#textorientation)|Gets or sets the text orientation of all the cells within the range.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatupdatedata#usestandardheight)|Determines if the row height of the Range object equals the standard height of the sheet.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatupdatedata#usestandardwidth)|Indicates whether the column width of the Range object equals the standard width of the sheet.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Represents the url target for the hyperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Represents the document reference target for the hyperlink.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#screentip)|Represents the string displayed when hovering over the hyperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Represents the string that is displayed in the top left most cell in the range.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hyperlink](/javascript/api/excel/excel.rangeloadoptions#hyperlink)|Represents the hyperlink for the current range.|
||[isEntireColumn](/javascript/api/excel/excel.rangeloadoptions#isentirecolumn)|Represents if the current range is an entire column. Read-only.|
||[isEntireRow](/javascript/api/excel/excel.rangeloadoptions#isentirerow)|Represents if the current range is an entire row. Read-only.|
||[numberFormatLocal](/javascript/api/excel/excel.rangeloadoptions#numberformatlocal)|Represents Excel's number format code for the given range as a string in the language of the user.|
||[style](/javascript/api/excel/excel.rangeloadoptions#style)|Represents the style of the current range.|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[hyperlink](/javascript/api/excel/excel.rangeupdatedata#hyperlink)|Represents the hyperlink for the current range.|
||[numberFormatLocal](/javascript/api/excel/excel.rangeupdatedata#numberformatlocal)|Represents Excel's number format code for the given range as a string in the language of the user.|
||[style](/javascript/api/excel/excel.rangeupdatedata#style)|Represents the style of the current range.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Deletes this style.|
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
||[set(properties: Excel.Style)](/javascript/api/excel/excel.style#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.StyleUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.style#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Represents the vertical alignment for the style. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Indicates if Microsoft Excel wraps the text in the object.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Adds a new style to the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Gets a style by name.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Gets the loaded child items in this collection.|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[$all](/javascript/api/excel/excel.stylecollectionloadoptions#$all)||
||[borders](/javascript/api/excel/excel.stylecollectionloadoptions#borders)|For EACH ITEM in the collection: A Border collection of four Border objects that represent the style of the four borders.|
||[builtIn](/javascript/api/excel/excel.stylecollectionloadoptions#builtin)|For EACH ITEM in the collection: Indicates if the style is a built-in style.|
||[fill](/javascript/api/excel/excel.stylecollectionloadoptions#fill)|For EACH ITEM in the collection: The Fill of the style.|
||[font](/javascript/api/excel/excel.stylecollectionloadoptions#font)|For EACH ITEM in the collection: A Font object that represents the font of the style.|
||[formulaHidden](/javascript/api/excel/excel.stylecollectionloadoptions#formulahidden)|For EACH ITEM in the collection: Indicates if the formula will be hidden when the worksheet is protected.|
||[horizontalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#horizontalalignment)|For EACH ITEM in the collection: Represents the horizontal alignment for the style. See Excel.HorizontalAlignment for details.|
||[includeAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#includealignment)|For EACH ITEM in the collection: Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|
||[includeBorder](/javascript/api/excel/excel.stylecollectionloadoptions#includeborder)|For EACH ITEM in the collection: Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|
||[includeFont](/javascript/api/excel/excel.stylecollectionloadoptions#includefont)|For EACH ITEM in the collection: Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|
||[includeNumber](/javascript/api/excel/excel.stylecollectionloadoptions#includenumber)|For EACH ITEM in the collection: Indicates if the style includes the NumberFormat property.|
||[includePatterns](/javascript/api/excel/excel.stylecollectionloadoptions#includepatterns)|For EACH ITEM in the collection: Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|
||[includeProtection](/javascript/api/excel/excel.stylecollectionloadoptions#includeprotection)|For EACH ITEM in the collection: Indicates if the style includes the FormulaHidden and Locked protection properties.|
||[indentLevel](/javascript/api/excel/excel.stylecollectionloadoptions#indentlevel)|For EACH ITEM in the collection: An integer from 0 to 250 that indicates the indent level for the style.|
||[locked](/javascript/api/excel/excel.stylecollectionloadoptions#locked)|For EACH ITEM in the collection: Indicates if the object is locked when the worksheet is protected.|
||[name](/javascript/api/excel/excel.stylecollectionloadoptions#name)|For EACH ITEM in the collection: The name of the style.|
||[numberFormat](/javascript/api/excel/excel.stylecollectionloadoptions#numberformat)|For EACH ITEM in the collection: The format code of the number format for the style.|
||[numberFormatLocal](/javascript/api/excel/excel.stylecollectionloadoptions#numberformatlocal)|For EACH ITEM in the collection: The localized format code of the number format for the style.|
||[readingOrder](/javascript/api/excel/excel.stylecollectionloadoptions#readingorder)|For EACH ITEM in the collection: The reading order for the style.|
||[shrinkToFit](/javascript/api/excel/excel.stylecollectionloadoptions#shrinktofit)|For EACH ITEM in the collection: Indicates if text automatically shrinks to fit in the available column width.|
||[verticalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#verticalalignment)|For EACH ITEM in the collection: Represents the vertical alignment for the style. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.stylecollectionloadoptions#wraptext)|For EACH ITEM in the collection: Indicates if Microsoft Excel wraps the text in the object.|
|[StyleData](/javascript/api/excel/excel.styledata)|[borders](/javascript/api/excel/excel.styledata#borders)|A Border collection of four Border objects that represent the style of the four borders.|
||[builtIn](/javascript/api/excel/excel.styledata#builtin)|Indicates if the style is a built-in style.|
||[fill](/javascript/api/excel/excel.styledata#fill)|The Fill of the style.|
||[font](/javascript/api/excel/excel.styledata#font)|A Font object that represents the font of the style.|
||[formulaHidden](/javascript/api/excel/excel.styledata#formulahidden)|Indicates if the formula will be hidden when the worksheet is protected.|
||[horizontalAlignment](/javascript/api/excel/excel.styledata#horizontalalignment)|Represents the horizontal alignment for the style. See Excel.HorizontalAlignment for details.|
||[includeAlignment](/javascript/api/excel/excel.styledata#includealignment)|Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|
||[includeBorder](/javascript/api/excel/excel.styledata#includeborder)|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|
||[includeFont](/javascript/api/excel/excel.styledata#includefont)|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|
||[includeNumber](/javascript/api/excel/excel.styledata#includenumber)|Indicates if the style includes the NumberFormat property.|
||[includePatterns](/javascript/api/excel/excel.styledata#includepatterns)|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|
||[includeProtection](/javascript/api/excel/excel.styledata#includeprotection)|Indicates if the style includes the FormulaHidden and Locked protection properties.|
||[indentLevel](/javascript/api/excel/excel.styledata#indentlevel)|An integer from 0 to 250 that indicates the indent level for the style.|
||[locked](/javascript/api/excel/excel.styledata#locked)|Indicates if the object is locked when the worksheet is protected.|
||[name](/javascript/api/excel/excel.styledata#name)|The name of the style.|
||[numberFormat](/javascript/api/excel/excel.styledata#numberformat)|The format code of the number format for the style.|
||[numberFormatLocal](/javascript/api/excel/excel.styledata#numberformatlocal)|The localized format code of the number format for the style.|
||[readingOrder](/javascript/api/excel/excel.styledata#readingorder)|The reading order for the style.|
||[shrinkToFit](/javascript/api/excel/excel.styledata#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
||[verticalAlignment](/javascript/api/excel/excel.styledata#verticalalignment)|Represents the vertical alignment for the style. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.styledata#wraptext)|Indicates if Microsoft Excel wraps the text in the object.|
|[StyleLoadOptions](/javascript/api/excel/excel.styleloadoptions)|[$all](/javascript/api/excel/excel.styleloadoptions#$all)||
||[borders](/javascript/api/excel/excel.styleloadoptions#borders)|A Border collection of four Border objects that represent the style of the four borders.|
||[builtIn](/javascript/api/excel/excel.styleloadoptions#builtin)|Indicates if the style is a built-in style.|
||[fill](/javascript/api/excel/excel.styleloadoptions#fill)|The Fill of the style.|
||[font](/javascript/api/excel/excel.styleloadoptions#font)|A Font object that represents the font of the style.|
||[formulaHidden](/javascript/api/excel/excel.styleloadoptions#formulahidden)|Indicates if the formula will be hidden when the worksheet is protected.|
||[horizontalAlignment](/javascript/api/excel/excel.styleloadoptions#horizontalalignment)|Represents the horizontal alignment for the style. See Excel.HorizontalAlignment for details.|
||[includeAlignment](/javascript/api/excel/excel.styleloadoptions#includealignment)|Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|
||[includeBorder](/javascript/api/excel/excel.styleloadoptions#includeborder)|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|
||[includeFont](/javascript/api/excel/excel.styleloadoptions#includefont)|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|
||[includeNumber](/javascript/api/excel/excel.styleloadoptions#includenumber)|Indicates if the style includes the NumberFormat property.|
||[includePatterns](/javascript/api/excel/excel.styleloadoptions#includepatterns)|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|
||[includeProtection](/javascript/api/excel/excel.styleloadoptions#includeprotection)|Indicates if the style includes the FormulaHidden and Locked protection properties.|
||[indentLevel](/javascript/api/excel/excel.styleloadoptions#indentlevel)|An integer from 0 to 250 that indicates the indent level for the style.|
||[locked](/javascript/api/excel/excel.styleloadoptions#locked)|Indicates if the object is locked when the worksheet is protected.|
||[name](/javascript/api/excel/excel.styleloadoptions#name)|The name of the style.|
||[numberFormat](/javascript/api/excel/excel.styleloadoptions#numberformat)|The format code of the number format for the style.|
||[numberFormatLocal](/javascript/api/excel/excel.styleloadoptions#numberformatlocal)|The localized format code of the number format for the style.|
||[readingOrder](/javascript/api/excel/excel.styleloadoptions#readingorder)|The reading order for the style.|
||[shrinkToFit](/javascript/api/excel/excel.styleloadoptions#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
||[verticalAlignment](/javascript/api/excel/excel.styleloadoptions#verticalalignment)|Represents the vertical alignment for the style. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.styleloadoptions#wraptext)|Indicates if Microsoft Excel wraps the text in the object.|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[borders](/javascript/api/excel/excel.styleupdatedata#borders)|A Border collection of four Border objects that represent the style of the four borders.|
||[fill](/javascript/api/excel/excel.styleupdatedata#fill)|The Fill of the style.|
||[font](/javascript/api/excel/excel.styleupdatedata#font)|A Font object that represents the font of the style.|
||[formulaHidden](/javascript/api/excel/excel.styleupdatedata#formulahidden)|Indicates if the formula will be hidden when the worksheet is protected.|
||[horizontalAlignment](/javascript/api/excel/excel.styleupdatedata#horizontalalignment)|Represents the horizontal alignment for the style. See Excel.HorizontalAlignment for details.|
||[includeAlignment](/javascript/api/excel/excel.styleupdatedata#includealignment)|Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|
||[includeBorder](/javascript/api/excel/excel.styleupdatedata#includeborder)|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|
||[includeFont](/javascript/api/excel/excel.styleupdatedata#includefont)|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|
||[includeNumber](/javascript/api/excel/excel.styleupdatedata#includenumber)|Indicates if the style includes the NumberFormat property.|
||[includePatterns](/javascript/api/excel/excel.styleupdatedata#includepatterns)|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|
||[includeProtection](/javascript/api/excel/excel.styleupdatedata#includeprotection)|Indicates if the style includes the FormulaHidden and Locked protection properties.|
||[indentLevel](/javascript/api/excel/excel.styleupdatedata#indentlevel)|An integer from 0 to 250 that indicates the indent level for the style.|
||[locked](/javascript/api/excel/excel.styleupdatedata#locked)|Indicates if the object is locked when the worksheet is protected.|
||[numberFormat](/javascript/api/excel/excel.styleupdatedata#numberformat)|The format code of the number format for the style.|
||[numberFormatLocal](/javascript/api/excel/excel.styleupdatedata#numberformatlocal)|The localized format code of the number format for the style.|
||[readingOrder](/javascript/api/excel/excel.styleupdatedata#readingorder)|The reading order for the style.|
||[shrinkToFit](/javascript/api/excel/excel.styleupdatedata#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
||[verticalAlignment](/javascript/api/excel/excel.styleupdatedata#verticalalignment)|Represents the vertical alignment for the style. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.styleupdatedata#wraptext)|Indicates if Microsoft Excel wraps the text in the object.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Occurs when data in cells changes on a specific table.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Occurs when the selection changes on a specific table.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Gets the address that represents the changed area of a table on a specific worksheet.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Gets the change type that represents how the Changed event is triggered. See Excel.DataChangeType for details.|
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
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[name](/javascript/api/excel/excel.workbookdata#name)|Gets the workbook name. Read-only.|
||[properties](/javascript/api/excel/excel.workbookdata#properties)|Gets the workbook properties. Read-only.|
||[protection](/javascript/api/excel/excel.workbookdata#protection)|Returns workbook protection object for a workbook. Read-only.|
||[styles](/javascript/api/excel/excel.workbookdata#styles)|Represents a collection of styles associated with the workbook. Read-only.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[name](/javascript/api/excel/excel.workbookloadoptions#name)|Gets the workbook name. Read-only.|
||[properties](/javascript/api/excel/excel.workbookloadoptions#properties)|Gets the workbook properties.|
||[protection](/javascript/api/excel/excel.workbookloadoptions#protection)|Returns workbook protection object for a workbook.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protects a workbook. Fails if the workbook has been protected.|
||[protected](/javascript/api/excel/excel.workbookprotection#protected)|Indicates if the workbook is protected. Read-Only.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Unprotects a workbook.|
|[WorkbookProtectionData](/javascript/api/excel/excel.workbookprotectiondata)|[protected](/javascript/api/excel/excel.workbookprotectiondata#protected)|Indicates if the workbook is protected. Read-Only.|
|[WorkbookProtectionLoadOptions](/javascript/api/excel/excel.workbookprotectionloadoptions)|[$all](/javascript/api/excel/excel.workbookprotectionloadoptions#$all)||
||[protected](/javascript/api/excel/excel.workbookprotectionloadoptions#protected)|Indicates if the workbook is protected. Read-Only.|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[properties](/javascript/api/excel/excel.workbookupdatedata#properties)|Gets the workbook properties.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Excel.Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copy a worksheet and place it at the specified position. Return the copied worksheet.|
||[copy(positionType?: Excel.WorksheetPositionType, relativeTo?: Excel.Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copy a worksheet and place it at the specified position. Return the copied worksheet.|
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
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Occurs when any worksheet in the workbook is activated.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Occurs when a new worksheet is added to the workbook.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Occurs when any worksheet in the workbook is deactivated.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Occurs when a worksheet is deleted from the workbook.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardheight)|For EACH ITEM in the collection: Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|
||[standardWidth](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardwidth)|For EACH ITEM in the collection: Returns or sets the standard (default) width of all the columns in the worksheet.|
||[tabColor](/javascript/api/excel/excel.worksheetcollectionloadoptions#tabcolor)|For EACH ITEM in the collection: Gets or sets the worksheet tab color.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[standardHeight](/javascript/api/excel/excel.worksheetdata#standardheight)|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|
||[standardWidth](/javascript/api/excel/excel.worksheetdata#standardwidth)|Returns or sets the standard (default) width of all the columns in the worksheet.|
||[tabColor](/javascript/api/excel/excel.worksheetdata#tabcolor)|Gets or sets the worksheet tab color.|
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
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetloadoptions#standardheight)|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|
||[standardWidth](/javascript/api/excel/excel.worksheetloadoptions#standardwidth)|Returns or sets the standard (default) width of all the columns in the worksheet.|
||[tabColor](/javascript/api/excel/excel.worksheetloadoptions#tabcolor)|Gets or sets the worksheet tab color.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Unprotects a worksheet.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Represents the worksheet protection option of allowing editing objects.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Represents the worksheet protection option of allowing editing scenarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Represents the worksheet protection option of selection mode.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Gets the range address that represents the selected area of a specific worksheet.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Gets the id of the worksheet in which the selection changed.|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[standardWidth](/javascript/api/excel/excel.worksheetupdatedata#standardwidth)|Returns or sets the standard (default) width of all the columns in the worksheet.|
||[tabColor](/javascript/api/excel/excel.worksheetupdatedata#tabcolor)|Gets or sets the worksheet tab color.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel&view=excel-js-1.7)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
