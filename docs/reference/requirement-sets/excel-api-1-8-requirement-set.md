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
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula1-member)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula2-member)|With the ternary operators Between and NotBetween, specifies the upper bound operand.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-operator-member)|The operator to use for validating the data.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#excel-excel-chart-categorylabellevel-member)|Specifies a chart category label level enumeration constant, referring to the level of the source category labels.|
||[displayBlanksAs](/javascript/api/excel/excel.chart#excel-excel-chart-displayblanksas-member)|Specifies the way that blank cells are plotted on a chart.|
||[onActivated](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member)|Occurs when the chart is activated.|
||[onDeactivated](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member)|Occurs when the chart is deactivated.|
||[plotArea](/javascript/api/excel/excel.chart#excel-excel-chart-plotarea-member)|Represents the plot area for the chart.|
||[plotBy](/javascript/api/excel/excel.chart#excel-excel-chart-plotby-member)|Specifies the way columns or rows are used as data series on the chart.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#excel-excel-chart-plotvisibleonly-member)|True if only visible cells are plotted.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#excel-excel-chart-seriesnamelevel-member)|Specifies a chart series name level enumeration constant, referring to the level of the source series names.|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#excel-excel-chart-showdatalabelsovermaximum-member)|Specifies whether to show the data labels when the value is greater than the maximum value on the value axis.|
||[style](/javascript/api/excel/excel.chart#excel-excel-chart-style-member)|Specifies the chart style for the chart.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-chartid-member)|Gets the ID of the chart that is activated.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is activated.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-chartid-member)|Gets the ID of the chart that is added to the worksheet.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is added.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[alignment](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-alignment-member)|Specifies the alignment for the specified axis tick label.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-isbetweencategories-member)|Specifies if the value axis crosses the category axis between categories.|
||[multiLevel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-multilevel-member)|Specifies if an axis is multilevel.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-numberformat-member)|Specifies the format code for the axis tick label.|
||[offset](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-offset-member)|Specifies the distance between the levels of labels, and the distance between the first level and the axis line.|
||[position](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-position-member)|Specifies the specified axis position where the other axis crosses.|
||[positionAt](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-positionat-member)|Specifies the axis position where the other axis crosses.|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setpositionat-member(1))|Sets the specified axis position where the other axis crosses.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-textorientation-member)|Specifies the angle to which the text is oriented for the chart axis tick label.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-fill-member)|Specifies chart fill formatting.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-setformula-member(1))|A string value that represents the formula of chart axis title using A1-style notation.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-border-member)|Specifies the chart axis title's border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-fill-member)|Specifies the chart axis title's fill formatting.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-clear-member(1))|Clear the border format of a chart element.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member)|Occurs when a chart is activated.|
||[onAdded](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member)|Occurs when a new chart is added to the worksheet.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member)|Occurs when a chart is deactivated.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member)|Occurs when a chart is deleted.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-autotext-member)|Specifies if the data label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-format-member)|Represents the format of chart data label.|
||[formula](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-formula-member)|String value that represents the formula of chart data label using A1-style notation.|
||[height](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-height-member)|Returns the height, in points, of the chart data label.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-horizontalalignment-member)|Represents the horizontal alignment for chart data label.|
||[left](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-left-member)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-numberformat-member)|String value that represents the format code for data label.|
||[text](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-text-member)|String representing the text of the data label on a chart.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-textorientation-member)|Represents the angle to which the text is oriented for the chart data label.|
||[top](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-top-member)|Represents the distance, in points, from the top edge of chart data label to the top of chart area.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-verticalalignment-member)|Represents the vertical alignment of chart data label.|
||[width](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-width-member)|Returns the width, in points, of the chart data label.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-border-member)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-autotext-member)|Specifies if data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-horizontalalignment-member)|Specifies the horizontal alignment for chart data label.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-numberformat-member)|Specifies the format code for data labels.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-textorientation-member)|Represents the angle to which the text is oriented for data labels.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-verticalalignment-member)|Represents the vertical alignment of chart data label.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-chartid-member)|Gets the ID of the chart that is deactivated.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is deactivated.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-chartid-member)|Gets the ID of the chart that is deleted from the worksheet.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is deleted.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-height-member)|Specifies the height of the legend entry on the chart legend.|
||[index](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-index-member)|Specifies the index of the legend entry in the chart legend.|
||[left](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-left-member)|Specifies the left value of a chart legend entry.|
||[top](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-top-member)|Specifies the top of a chart legend entry.|
||[width](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-width-member)|Represents the width of the legend entry on the chart Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-border-member)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[format](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-format-member)|Specifies the formatting of a chart plot area.|
||[height](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-height-member)|Specifies the height value of a plot area.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideheight-member)|Specifies the inside height value of a plot area.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideleft-member)|Specifies the inside left value of a plot area.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidetop-member)|Specifies the inside top value of a plot area.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidewidth-member)|Specifies the inside width value of a plot area.|
||[left](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-left-member)|Specifies the left value of a plot area.|
||[position](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-position-member)|Specifies the position of a plot area.|
||[top](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-top-member)|Specifies the top value of a plot area.|
||[width](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-width-member)|Specifies the width value of a plot area.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-border-member)|Specifies the border attributes of a chart plot area.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-fill-member)|Specifies the fill format of an object, which includes background formatting information.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-axisgroup-member)|Specifies the group for the specified series.|
||[dataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-datalabels-member)|Represents a collection of all data labels in the series.|
||[explosion](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-explosion-member)|Specifies the explosion value for a pie-chart or doughnut-chart slice.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-firstsliceangle-member)|Specifies the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertifnegative-member)|True if Excel inverts the pattern in the item when it corresponds to a negative number.|
||[overlap](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-overlap-member)|Specifies how bars and columns are positioned.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-secondplotsize-member)|Specifies the size of the secondary section of either a pie-of-pie chart or a bar-of-pie chart, as a percentage of the size of the primary pie.|
||[splitType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splittype-member)|Specifies the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-varybycategories-member)|True if Excel assigns a different color or pattern to each data marker.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-backwardperiod-member)|Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-forwardperiod-member)|Represents the number of periods that the trendline extends forward.|
||[label](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-label-member)|Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showequation-member)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showrsquared-member)|True if the r-squared value for the trendline is displayed on the chart.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-autotext-member)|Specifies if the trendline label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-format-member)|The format of the chart trendline label.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-formula-member)|String value that represents the formula of the chart trendline label using A1-style notation.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-height-member)|Returns the height, in points, of the chart trendline label.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-horizontalalignment-member)|Represents the horizontal alignment of the chart trendline label.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-left-member)|Represents the distance, in points, from the left edge of the chart trendline label to the left edge of the chart area.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-numberformat-member)|String value that represents the format code for the trendline label.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-text-member)|String representing the text of the trendline label on a chart.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-textorientation-member)|Represents the angle to which the text is oriented for the chart trendline label.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-top-member)|Represents the distance, in points, from the top edge of the chart trendline label to the top of the chart area.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-verticalalignment-member)|Represents the vertical alignment of the chart trendline label.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-width-member)|Returns the width, in points, of the chart trendline label.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-border-member)|Specifies the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-fill-member)|Specifies the fill format of the current chart trendline label.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-font-member)|Specifies the font attributes (such as font name, font size, and color) for a chart trendline label.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#excel-excel-customdatavalidation-formula-member)|A custom data validation formula.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[field](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-field-member)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-id-member)|ID of the DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-name-member)|Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-numberformat-member)|Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-position-member)|Position of the DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-settodefault-member(1))|Reset the DataPivotHierarchy back to its default values.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-showas-member)|Specifies if the data should be shown as a specific summary calculation.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-summarizeby-member)|Specifies if all items of the DataPivotHierarchy are shown.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitem-member(1))|Gets a DataPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitemornullobject-member(1))|Gets a DataPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-clear-member(1))|Clears the data validation from the current range.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-erroralert-member)|Error alert when user enters invalid data.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-ignoreblanks-member)|Specifies if data validation will be performed on blank cells.|
||[prompt](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-prompt-member)|Prompt when users select a cell.|
||[rule](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-rule-member)|Data validation rule that contains different type of data validation criteria.|
||[type](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-type-member)|Type of the data validation, see `Excel.DataValidationType` for details.|
||[valid](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-valid-member)|Represents if all cell values are valid according to the data validation rules.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-message-member)|Represents the error alert message.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-showalert-member)|Specifies whether to show an error alert dialog when a user enters invalid data.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-style-member)|The data validation alert type, please see `Excel.DataValidationAlertStyle` for details.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-title-member)|Represents the error alert dialog title.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-message-member)|Specifies the message of the prompt.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-showprompt-member)|Specifies if a prompt is shown when a user selects a cell with data validation.|
||[title](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-title-member)|Specifies the title for the prompt.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[custom](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-custom-member)|Custom data validation criteria.|
||[date](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-date-member)|Date data validation criteria.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-decimal-member)|Decimal data validation criteria.|
||[list](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-list-member)|List data validation criteria.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-textlength-member)|Text length data validation criteria.|
||[time](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-time-member)|Time data validation criteria.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-wholenumber-member)|Whole number data validation criteria.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula1-member)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula2-member)|With the ternary operators Between and NotBetween, specifies the upper bound operand.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-operator-member)|The operator to use for validating the data.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-enablemultiplefilteritems-member)|Determines whether to allow multiple filter items.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-fields-member)|Returns the PivotFields associated with the FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-id-member)|ID of the FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-name-member)|Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-position-member)|Position of the FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-settodefault-member(1))|Reset the FilterPivotHierarchy back to its default values.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitem-member(1))|Gets a FilterPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitemornullobject-member(1))|Gets a FilterPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-incelldropdown-member)|Specifies whether to display the list in a cell drop-down.|
||[source](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-source-member)|Source of the list for data validation|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[id](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-id-member)|ID of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-items-member)|Returns the PivotItems associated with the PivotField.|
||[name](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-name-member)|Name of the PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-showallitems-member)|Determines whether to show all items of the PivotField.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbylabels-member(1))|Sorts the PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-subtotals-member)|Subtotals of the PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getcount-member(1))|Gets the number of pivot fields in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitem-member(1))|Gets a PivotField by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitemornullobject-member(1))|Gets a PivotField by name.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[fields](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-fields-member)|Returns the PivotFields associated with the PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-id-member)|ID of the PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-name-member)|Name of the PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitem-member(1))|Gets a PivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitemornullobject-member(1))|Gets a PivotHierarchy by name.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[id](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-id-member)|ID of the PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-isexpanded-member)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-name-member)|Name of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-visible-member)|Specifies if the PivotItem is visible.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getcount-member(1))|Gets the number of PivotItems in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitem-member(1))|Gets a PivotItem by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitemornullobject-member(1))|Gets a PivotItem by name.|
||[items](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcolumnlabelrange-member(1))|Returns the range where the PivotTable's column labels reside.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatabodyrange-member(1))|Returns the range where the PivotTable's data values reside.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getfilteraxisrange-member(1))|Returns the range of the PivotTable's filter area.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrange-member(1))|Returns the range the PivotTable exists on, excluding the filter area.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrowlabelrange-member(1))|Returns the range where the PivotTable's row labels reside.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-layouttype-member)|This property indicates the PivotLayoutType of all fields on the PivotTable.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showcolumngrandtotals-member)|Specifies if the PivotTable report shows grand totals for columns.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showrowgrandtotals-member)|Specifies if the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-subtotallocation-member)|This property indicates the `SubtotalLocationType` of all fields on the PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[columnHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-columnhierarchies-member)|The Column Pivot Hierarchies of the PivotTable.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-datahierarchies-member)|The Data Pivot Hierarchies of the PivotTable.|
||[delete()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-delete-member(1))|Deletes the PivotTable.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-filterhierarchies-member)|The Filter Pivot Hierarchies of the PivotTable.|
||[hierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-hierarchies-member)|The Pivot Hierarchies of the PivotTable.|
||[layout](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-layout-member)|The PivotLayout describing the layout and visual structure of the PivotTable.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-rowhierarchies-member)|The Row Pivot Hierarchies of the PivotTable.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-add-member(1))|Add a PivotTable based on the specified source data and insert it at the top-left cell of the destination range.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#excel-excel-range-datavalidation-member)|Returns a data validation object.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-fields-member)|Returns the PivotFields associated with the RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-id-member)|ID of the RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-name-member)|Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-position-member)|Position of the RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-settodefault-member(1))|Reset the RowColumnPivotHierarchy back to its default values.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitem-member(1))|Gets a RowColumnPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitemornullobject-member(1))|Gets a RowColumnPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#excel-excel-runtime-enableevents-member)|Toggle JavaScript events in the current task pane or content add-in.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-basefield-member)|The PivotField to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.|
||[baseItem](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-baseitem-member)|The item to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.|
||[calculation](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-calculation-member)|The `ShowAs` calculation to use for the PivotField.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#excel-excel-style-autoindent-member)|Specifies if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/javascript/api/excel/excel.style#excel-excel-style-textorientation-member)|The text orientation for the style.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-automatic-member)|If `Automatic` is set to `true`, then all other values will be ignored when setting the `Subtotals`.|
||[average](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-average-member)||
||[count](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-count-member)||
||[countNumbers](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-countnumbers-member)||
||[max](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-max-member)||
||[min](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-min-member)||
||[product](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-product-member)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviation-member)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviationp-member)||
||[sum](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-sum-member)||
||[variance](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variance-member)||
||[varianceP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variancep-member)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#excel-excel-table-legacyid-member)|Returns a numeric ID.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrange-member(1))|Gets the range that represents the changed area of a table on a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrangeornullobject-member(1))|Gets the range that represents the changed area of a table on a specific worksheet.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#excel-excel-workbook-readonly-member)|Returns `true` if the workbook is open in read-only mode.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member)|Occurs when the worksheet is calculated.|
||[showGridlines](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showgridlines-member)|Specifies if gridlines are visible to the user.|
||[showHeadings](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showheadings-member)|Specifies if headings are visible to the user.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the calculation occurred.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrange-member(1))|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrangeornullobject-member(1))|Gets the range that represents the changed area of a specific worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member)|Occurs when any worksheet in the workbook is calculated.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
