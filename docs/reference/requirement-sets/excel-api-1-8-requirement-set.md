---
title: Excel JavaScript API requirement set 1.8
description: 'Details about the ExcelApi 1.8 requirement set'
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
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

The following table lists the APIs in Excel JavaScript API requirement set 1.8. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.8 or earlier, see [Excel APIs in requirement set 1.8 or earlier](/javascript/api/excel?view=excel-js-1.8).

| Class | Fields | Description |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand.|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|The operator to use for validating the data.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Returns or sets a ChartCategoryLabelLevel enumeration constant referring to|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Returns or sets the way that blank cells are plotted on a chart. Read/Write.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Returns or sets the way columns or rows are used as data series on the chart. Read/Write.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|True if only visible cells are plotted. False if both visible and hidden cells are plotted. Read/Write.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Occurs when the chart is activated.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Occurs when the chart is deactivated.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Represents the plotArea for the chart.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Returns or sets a ChartSeriesNameLevel enumeration constant referring to|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Represents whether to show the data labels when the value is greater than the maximum value on the value axis.|
||[style](/javascript/api/excel/excel.chart#style)|Returns or sets the chart style for the chart. Read/Write.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Gets the id of the chart that is activated.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Gets the id of the worksheet in which the chart is activated.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Gets the id of the chart that is added to the worksheet.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Gets the id of the worksheet in which the chart is added.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[alignment](/javascript/api/excel/excel.chartaxis#alignment)|Represents the alignment for the specified axis tick label. See Excel.ChartTextHorizontalAlignment for detail.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Represents whether value axis crosses the category axis between categories.|
||[multiLevel](/javascript/api/excel/excel.chartaxis#multilevel)|Represents whether an axis is multilevel or not.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Represents the format code for the axis tick label.|
||[offset](/javascript/api/excel/excel.chartaxis#offset)|Represents the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Represents the specified axis position where the other axis crosses. See Excel.ChartAxisPosition for details.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|Represents the specified axis position where the other axis crosses at. You should use the SetPositionAt(double) method to set this property.|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Set the specified axis position where the other axis crosses at.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Represents the text orientation of the axis tick label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Represents chart fill formatting. Read-only.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|A string value that represents the formula of chart axis title using A1-style notation.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#border)|Represents the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Represents chart fill formatting.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Clear the border format of a chart element.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Occurs when a chart is activated.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Occurs when a new chart is added to the worksheet.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Occurs when a chart is deactivated.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Occurs when a chart is deleted.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autotext)|Boolean value representing if data label automatically generates appropriate text based on context.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|String value that represents the formula of chart data label using A1-style notation.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|String value that represents the format code for data label.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Represents the format of chart data label.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|String representing the text of the data label on a chart.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#border)|Represents the border format, which includes color, linestyle, and weight. Read-only.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autotext)|Represents whether data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Represents the format code for data labels.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Gets the id of the chart that is deactivated.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Gets the id of the worksheet in which the chart is deactivated.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Gets the id of the chart that is deleted from the worksheet.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Gets the id of the worksheet in which the chart is deleted.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Represents the height of the legendEntry on the chart legend.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Represents the index of the legendEntry in the chart legend.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Represents the left of a chart legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Represents the top of a chart legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Represents the width of the legendEntry on the chart Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#border)|Represents the border format, which includes color, linestyle, and weight. Read-only.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Represents the height value of plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Represents the insideHeight value of plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Represents the insideLeft value of plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Represents the insideTop value of plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Represents the insideWidth value of plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Represents the left value of plotArea.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Represents the position of plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Represents the formatting of a chart plotArea.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Represents the top value of plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Represents the width value of plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#border)|Represents the border attributes of a chart plotArea.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Represents the fill format of an object, which includes background formatting information.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Returns or sets the group for the specified series. Read/Write|
||[explosion](/javascript/api/excel/excel.chartseries#explosion)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/Write.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/Write|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/Write.|
||[overlap](/javascript/api/excel/excel.chartseries#overlap)|Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts. Read/Write.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Represents a collection of all dataLabels in the series.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/Write.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/Write.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/Write.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Represents the number of periods that the trendline extends forward.|
||[label](/javascript/api/excel/excel.charttrendline#label)|Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|True if the R-squared for the trendline is displayed on the chart.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Boolean value representing if trendline label automatically generates appropriate text based on context.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|String value that represents the formula of chart trendline label using A1-style notation.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Represents the horizontal alignment for chart trendline label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|String value that represents the format code for trendline label.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Represents the format of chart trendline label.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|String representing the text of the trendline label on a chart.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Represents the vertical alignment of chart trendline label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Represents the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Represents the fill format of the current chart trendline label.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Represents the font attributes (font name, font size, color, etc.) for a chart trendline label.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|A custom data validation formula. This creates special input rules, such as preventing duplicates, or limiting the total in a range of cells.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Position of the DataPivotHierarchy.|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|Id of the DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Reset the DataPivotHierarchy back to its default values.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Determines whether the data should be shown as a specific summary calculation or not.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Determines whether to show all items of the DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Gets a DataPivotHierarchy by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Gets a DataPivotHierarchy by name. If the DataPivotHierarchy does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Removes the PivotHierarchy from the current axis.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Clears the data validation from the current range.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Error alert when user enters invalid data.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Prompt when users select a cell.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Type of the data validation, see Excel.DataValidationType for details.|
||[valid](/javascript/api/excel/excel.datavalidation#valid)|Represents if all cell values are valid according to the data validation rules.|
||[rule](/javascript/api/excel/excel.datavalidation#rule)|Data validation rule that contains different type of data validation criteria.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Represents error alert message.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Determines whether to show an error alert dialog or not when a user enters invalid data. The default is true.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Represents data validation alert type, please see Excel.DataValidationAlertStyle for details.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Represents error alert dialog title.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Represents the message of the prompt.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Determines whether or not to show the prompt when user selects a cell with data validation.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Represents the title for the prompt.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[custom](/javascript/api/excel/excel.datavalidationrule#custom)|Custom data validation criteria.|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|Date data validation criteria.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|Decimal data validation criteria.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|List data validation criteria.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|TextLength data validation criteria.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Time data validation criteria.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|WholeNumber data validation criteria.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand.|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|The operator to use for validating the data.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Determines whether to allow multiple filter items.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Position of the FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Returns the PivotFields associated with the FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|Id of the FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Reset the FilterPivotHierarchy back to its default values.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Gets a FilterPivotHierarchy by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Gets a FilterPivotHierarchy by name. If the FilterPivotHierarchy does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Removes the PivotHierarchy from the current axis.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Displays the list in cell drop down or not, it defaults to true.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Source of the list for data validation|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Name of the PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|Id of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Returns the PivotItems that comprise with the PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Determines whether to show all items of the PivotField.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Subtotals of the PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Gets the number of pivot fields in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Gets a PivotField by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Gets a PivotField by name. If the PivotField does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Gets the loaded child items in this collection.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Name of the PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Returns the PivotFields associated with the PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|Id of the PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Gets a PivotHierarchy by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Gets the loaded child items in this collection.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Name of the PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|Id of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Determines whether the PivotItem is visible or not.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Gets the number of pivot items in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Gets a PivotItem by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Gets a PivotItem by name. If the PivotItem does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Gets the loaded child items in this collection.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Returns the range where the PivotTable's column labels reside.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Returns the range where the PivotTable's data values reside.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Returns the range of the PivotTable's filter area.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Returns the range the PivotTable exists on, excluding the filter area.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Returns the range where the PivotTable's row labels reside.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Specifies whether the PivotTable report shows grand totals for columns.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Specifies whether the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Deletes the PivotTable.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|The Column Pivot Hierarchies of the PivotTable.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|The Data Pivot Hierarchies of the PivotTable.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|The Filter Pivot Hierarchies of the PivotTable.|
||[hierarchies](/javascript/api/excel/excel.pivottable#hierarchies)|The Pivot Hierarchies of the PivotTable.|
||[layout](/javascript/api/excel/excel.pivottable#layout)|The PivotLayout describing the layout and visual structure of the PivotTable.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|The Row Pivot Hierarchies of the PivotTable.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Add a Pivottable based on the specified source data and insert it at the top left cell of the destination range.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Returns a data validation object.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Position of the RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Returns the PivotFields associated with the RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|Id of the RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Reset the RowColumnPivotHierarchy back to its default values.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Gets a RowColumnPivotHierarchy by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Removes the PivotHierarchy from the current axis.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Toggle JavaScript events in the current task pane or content add-in.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|The base PivotField to base the ShowAs calculation, if applicable based on the ShowAsCalculation type, else null.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|The base Item to base the ShowAs calculation on, if applicable based on the ShowAsCalculation type, else null.|
||[calculation](/javascript/api/excel/excel.showasrule#calculation)|The ShowAs Calculation to use for the Data PivotField. See Excel.ShowAsCalculation for Details.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|The text orientation for the style.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|
||[average](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[product](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[sum](/javascript/api/excel/excel.subtotals#sum)||
||[variance](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|Returns a numeric id.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Gets the range that represents the changed area of a table on a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Gets the range that represents the changed area of a table on a specific worksheet. It might return null object.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|True if the workbook is open in Read-only mode. Read-only.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|Occurs when the worksheet is calculated.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|Gets or sets the worksheet's gridlines flag.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|Gets or sets the worksheet's headings flag.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Gets the id of the worksheet that is calculated.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Gets the range that represents the changed area of a specific worksheet. It might return null object.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Occurs when any worksheet in the workbook is calculated.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.8)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
