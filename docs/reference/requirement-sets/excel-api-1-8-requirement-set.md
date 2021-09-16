---
title: Excel JavaScript API requirement set 1.8
description: 'Details about the ExcelApi 1.8 requirement set.'
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.8

The Excel JavaScript API requirement set 1.8 features include APIs for PivotTables, data validation, charts, events for charts, performance options, and workbook creation.

## PivotTable

Wave 2 of the PivotTable APIs lets add-ins set the hierarchies of a PivotTable. You can now control the data and how it is aggregated. Our [PivotTable article](../../excel/excel-add-ins-pivottables.md) has more on the new PivotTable functionality.

## Data Validation

Data validation gives you control of what a user enters in a worksheet. You can limit cells to pre-defined answer sets or give pop-up warnings about undesirable input. Learn more about [adding data validation to ranges](../../excel/excel-add-ins-data-validation.md) today.

## Charts

Another round of Chart APIs brings even greater programmatic control over chart elements. You now have greater access to the legend, axes, trendline, and plot area.

## Events

More [events](../../excel/excel-add-ins-events.md) have been added for charts. Have your add-in react to users interacting with the chart. You can also [toggle events](../../excel/performance.md#enable-and-disable-events) firing across the entire workbook.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.8. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.8 or earlier, see [Excel APIs in requirement set 1.8 or earlier](/javascript/api/excel?view=excel-js-1.8&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|With the ternary operators Between and NotBetween, specifies the upper bound operand.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|The operator to use for validating the data.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categoryLabelLevel)|Specifies a chart category label level enumeration constant, referring to the level of the source category labels.|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayBlanksAs)|Specifies the way that blank cells are plotted on a chart.|
||[plotBy](/javascript/api/excel/excel.chart#plotBy)|Specifies the way columns or rows are used as data series on the chart.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotVisibleOnly)|True if only visible cells are plotted.|
||[onActivated](/javascript/api/excel/excel.chart#onActivated)|Occurs when the chart is activated.|
||[onDeactivated](/javascript/api/excel/excel.chart#onDeactivated)|Occurs when the chart is deactivated.|
||[plotArea](/javascript/api/excel/excel.chart#plotArea)|Represents the plot area for the chart.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesNameLevel)|Specifies a chart series name level enumeration constant, referring to the level of the source series names.|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showDataLabelsOverMaximum)|Specifies whether to show the data labels when the value is greater than the maximum value on the value axis.|
||[style](/javascript/api/excel/excel.chart#style)|Specifies the chart style for the chart.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartId)|Gets the ID of the chart that is activated.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetId)|Gets the ID of the worksheet in which the chart is activated.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartId)|Gets the ID of the chart that is added to the worksheet.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetId)|Gets the ID of the worksheet in which the chart is added.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[alignment](/javascript/api/excel/excel.chartaxis#alignment)|Specifies the alignment for the specified axis tick label.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isBetweenCategories)|Specifies if the value axis crosses the category axis between categories.|
||[multiLevel](/javascript/api/excel/excel.chartaxis#multiLevel)|Specifies if an axis is multilevel.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberFormat)|Specifies the format code for the axis tick label.|
||[offset](/javascript/api/excel/excel.chartaxis#offset)|Specifies the distance between the levels of labels, and the distance between the first level and the axis line.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Specifies the specified axis position where the other axis crosses.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionAt)|Specifies the axis position where the other axis crosses.|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#setPositionAt_value_)|Sets the specified axis position where the other axis crosses.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textOrientation)|Specifies the angle to which the text is oriented for the chart axis tick label.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Specifies chart fill formatting.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setFormula_formula_)|A string value that represents the formula of chart axis title using A1-style notation.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#border)|Specifies the chart axis title's border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Specifies the chart axis title's fill formatting.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear__)|Clear the border format of a chart element.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onActivated)|Occurs when a chart is activated.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onAdded)|Occurs when a new chart is added to the worksheet.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#onDeactivated)|Occurs when a chart is deactivated.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#onDeleted)|Occurs when a chart is deleted.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autoText)|Specifies if the data label automatically generates appropriate text based on context.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|String value that represents the formula of chart data label using A1-style notation.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalAlignment)|Represents the horizontal alignment for chart data label.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberFormat)|String value that represents the format code for data label.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Represents the format of chart data label.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Returns the height, in points, of the chart data label.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Returns the width, in points, of the chart data label.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|String representing the text of the data label on a chart.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textOrientation)|Represents the angle to which the text is oriented for the chart data label.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Represents the distance, in points, from the top edge of chart data label to the top of chart area.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalAlignment)|Represents the vertical alignment of chart data label.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autoText)|Specifies if data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalAlignment)|Specifies the horizontal alignment for chart data label.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberFormat)|Specifies the format code for data labels.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textOrientation)|Represents the angle to which the text is oriented for data labels.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalAlignment)|Represents the vertical alignment of chart data label.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartId)|Gets the ID of the chart that is deactivated.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetId)|Gets the ID of the worksheet in which the chart is deactivated.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartId)|Gets the ID of the chart that is deleted from the worksheet.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetId)|Gets the ID of the worksheet in which the chart is deleted.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Specifies the height of the legend entry on the chart legend.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Specifies the index of the legend entry in the chart legend.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Specifies the left value of a chart legend entry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Specifies the top of a chart legend entry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Represents the width of the legend entry on the chart Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Specifies the height value of a plot area.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideHeight)|Specifies the inside height value of a plot area.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideLeft)|Specifies the inside left value of a plot area.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insideTop)|Specifies the inside top value of a plot area.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insideWidth)|Specifies the inside width value of a plot area.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Specifies the left value of a plot area.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Specifies the position of a plot area.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Specifies the formatting of a chart plot area.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Specifies the top value of a plot area.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Specifies the width value of a plot area.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#border)|Specifies the border attributes of a chart plot area.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Specifies the fill format of an object, which includes background formatting information.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisGroup)|Specifies the group for the specified series.|
||[explosion](/javascript/api/excel/excel.chartseries#explosion)|Specifies the explosion value for a pie-chart or doughnut-chart slice.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstSliceAngle)|Specifies the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertIfNegative)|True if Excel inverts the pattern in the item when it corresponds to a negative number.|
||[overlap](/javascript/api/excel/excel.chartseries#overlap)|Specifies how bars and columns are positioned.|
||[dataLabels](/javascript/api/excel/excel.chartseries#dataLabels)|Represents a collection of all data labels in the series.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondPlotSize)|Specifies the size of the secondary section of either a pie-of-pie chart or a bar-of-pie chart, as a percentage of the size of the primary pie.|
||[splitType](/javascript/api/excel/excel.chartseries#splitType)|Specifies the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varyByCategories)|True if Excel assigns a different color or pattern to each data marker.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardPeriod)|Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardPeriod)|Represents the number of periods that the trendline extends forward.|
||[label](/javascript/api/excel/excel.charttrendline#label)|Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showEquation)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showRSquared)|True if the r-squared value for the trendline is displayed on the chart.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autoText)|Specifies if the trendline label automatically generates appropriate text based on context.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|String value that represents the formula of the chart trendline label using A1-style notation.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalAlignment)|Represents the horizontal alignment of the chart trendline label.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Represents the distance, in points, from the left edge of the chart trendline label to the left edge of the chart area.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberFormat)|String value that represents the format code for the trendline label.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|The format of the chart trendline label.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Returns the height, in points, of the chart trendline label.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Returns the width, in points, of the chart trendline label.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|String representing the text of the trendline label on a chart.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textOrientation)|Represents the angle to which the text is oriented for the chart trendline label.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Represents the distance, in points, from the top edge of the chart trendline label to the top of the chart area.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalAlignment)|Represents the vertical alignment of the chart trendline label.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Specifies the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Specifies the fill format of the current chart trendline label.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Specifies the font attributes (such as font name, font size, and color) for a chart trendline label.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|A custom data validation formula.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberFormat)|Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Position of the DataPivotHierarchy.|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID of the DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#setToDefault__)|Reset the DataPivotHierarchy back to its default values.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showAs)|Specifies if the data should be shown as a specific summary calculation.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeBy)|Specifies if all items of the DataPivotHierarchy are shown.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add_pivotHierarchy_)|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getCount__)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getItem_name_)|Gets a DataPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getItemOrNullObject_name_)|Gets a DataPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove_DataPivotHierarchy_)|Removes the PivotHierarchy from the current axis.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear__)|Clears the data validation from the current range.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#errorAlert)|Error alert when user enters invalid data.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreBlanks)|Specifies if data validation will be performed on blank cells.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Prompt when users select a cell.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Type of the data validation, see `Excel.DataValidationType` for details.|
||[valid](/javascript/api/excel/excel.datavalidation#valid)|Represents if all cell values are valid according to the data validation rules.|
||[rule](/javascript/api/excel/excel.datavalidation#rule)|Data validation rule that contains different type of data validation criteria.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Represents the error alert message.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showAlert)|Specifies whether to show an error alert dialog when a user enters invalid data.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|The data validation alert type, please see `Excel.DataValidationAlertStyle` for details.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Represents the error alert dialog title.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Specifies the message of the prompt.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showPrompt)|Specifies if a prompt is shown when a user selects a cell with data validation.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Specifies the title for the prompt.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[custom](/javascript/api/excel/excel.datavalidationrule#custom)|Custom data validation criteria.|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|Date data validation criteria.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|Decimal data validation criteria.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|List data validation criteria.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textLength)|Text length data validation criteria.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Time data validation criteria.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholeNumber)|Whole number data validation criteria.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|With the ternary operators Between and NotBetween, specifies the upper bound operand.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|The operator to use for validating the data.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enableMultipleFilterItems)|Determines whether to allow multiple filter items.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Position of the FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Returns the PivotFields associated with the FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID of the FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#setToDefault__)|Reset the FilterPivotHierarchy back to its default values.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add_pivotHierarchy_)|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getCount__)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getItem_name_)|Gets a FilterPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getItemOrNullObject_name_)|Gets a FilterPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove_filterPivotHierarchy_)|Removes the PivotHierarchy from the current axis.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#inCellDropDown)|Specifies whether to display the list in a cell drop-down.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Source of the list for data validation|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Name of the PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Returns the PivotFields associated with the PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showAllItems)|Determines whether to show all items of the PivotField.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortByLabels_sortBy_)|Sorts the PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Subtotals of the PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getCount__)|Gets the number of pivot fields in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getItem_name_)|Gets a PivotField by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getItemOrNullObject_name_)|Gets a PivotField by name.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Gets the loaded child items in this collection.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Name of the PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Returns the PivotFields associated with the PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID of the PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getCount__)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getItem_name_)|Gets a PivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getItemOrNullObject_name_)|Gets a PivotHierarchy by name.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Gets the loaded child items in this collection.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isExpanded)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Name of the PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Specifies if the PivotItem is visible.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getCount__)|Gets the number of PivotItems in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getItem_name_)|Gets a PivotItem by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getItemOrNullObject_name_)|Gets a PivotItem by name.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Gets the loaded child items in this collection.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getColumnLabelRange__)|Returns the range where the PivotTable's column labels reside.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getDataBodyRange__)|Returns the range where the PivotTable's data values reside.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getFilterAxisRange__)|Returns the range of the PivotTable's filter area.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getRange__)|Returns the range the PivotTable exists on, excluding the filter area.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getRowLabelRange__)|Returns the range where the PivotTable's row labels reside.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layoutType)|This property indicates the PivotLayoutType of all fields on the PivotTable.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showColumnGrandTotals)|Specifies if the PivotTable report shows grand totals for columns.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showRowGrandTotals)|Specifies if the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotalLocation)|This property indicates the `SubtotalLocationType` of all fields on the PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete__)|Deletes the PivotTable.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnHierarchies)|The Column Pivot Hierarchies of the PivotTable.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#dataHierarchies)|The Data Pivot Hierarchies of the PivotTable.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterHierarchies)|The Filter Pivot Hierarchies of the PivotTable.|
||[hierarchies](/javascript/api/excel/excel.pivottable#hierarchies)|The Pivot Hierarchies of the PivotTable.|
||[layout](/javascript/api/excel/excel.pivottable#layout)|The PivotLayout describing the layout and visual structure of the PivotTable.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowHierarchies)|The Row Pivot Hierarchies of the PivotTable.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add_name__source__destination_)|Add a PivotTable based on the specified source data and insert it at the top-left cell of the destination range.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#dataValidation)|Returns a data validation object.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Position of the RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Returns the PivotFields associated with the RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID of the RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#setToDefault__)|Reset the RowColumnPivotHierarchy back to its default values.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add_pivotHierarchy_)|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getCount__)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getItem_name_)|Gets a RowColumnPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getItemOrNullObject_name_)|Gets a RowColumnPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove_rowColumnPivotHierarchy_)|Removes the PivotHierarchy from the current axis.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableEvents)|Toggle JavaScript events in the current task pane or content add-in.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#baseField)|The PivotField to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseItem)|The item to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.|
||[calculation](/javascript/api/excel/excel.showasrule#calculation)|The `ShowAs` calculation to use for the PivotField.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoIndent)|Specifies if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/javascript/api/excel/excel.style#textOrientation)|The text orientation for the style.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|If `Automatic` is set to `true`, then all other values will be ignored when setting the `Subtotals`.|
||[average](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countNumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[product](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standardDeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standardDeviationP)||
||[sum](/javascript/api/excel/excel.subtotals#sum)||
||[variance](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#varianceP)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyId)|Returns a numeric ID.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getRange_ctx_)|Gets the range that represents the changed area of a table on a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getRangeOrNullObject_ctx_)|Gets the range that represents the changed area of a table on a specific worksheet.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readOnly)|Returns `true` if the workbook is open in read-only mode.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#onCalculated)|Occurs when the worksheet is calculated.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showGridlines)|Specifies if gridlines are visible to the user.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showHeadings)|Specifies if headings are visible to the user.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetId)|Gets the ID of the worksheet in which the calculation occurred.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getRange_ctx_)|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getRangeOrNullObject_ctx_)|Gets the range that represents the changed area of a specific worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#onCalculated)|Occurs when any worksheet in the workbook is calculated.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
