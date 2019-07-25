---
title: Excel JavaScript API requirement set 1.8
description: 'Details about the ExcelApi 1.8 requirement set'
ms.date: 07/25/2019
ms.prod: excel
localization_priority: Normal
---

# What’s new in Excel JavaScript API 1.8

The Excel JavaScript API requirement set 1.8 features include APIs for PivotTables, data validation, charts, events for charts, performance options, and workbook creation.

## PivotTable

Wave 2 of the PivotTable APIs lets add-ins set the hierarchies of a PivotTable. You can now control the data and how it is aggregated. Our [PivotTable article](/office/dev/add-ins/excel/excel-add-ins-pivottables) has more on the new PivotTable functionality.

## Data Validation

Data validation gives you control of what a user enters in a worksheet. You can limit cells to pre-defined answer sets or give pop-up warnings about undesirable input. Learn more about [adding data validation to ranges](/office/dev/add-ins/excel/excel-add-ins-data-validation) today.

## Charts

Another round of Chart APIs brings even greater programmatic control over chart elements. You now have greater access to the legend, axes, trendline, and plot area.

## Events

More [events](/office/dev/add-ins/excel/excel-add-ins-events) have been added for charts. Have your add-in react to users interacting with the chart. You can also [toggle events](/office/dev/add-ins/excel/performance#enable-and-disable-events) firing across the entire workbook.

## API list

To see a complete list of all APIs supported by this requirement set (including previously released APIs), [click here to see a version-specific of the API reference documentation]((/javascript/api/excel?view=excel-js-1.8)).

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
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[alignment](/javascript/api/excel/excel.chartaxisdata#alignment)|Represents the alignment for the specified axis tick label. See Excel.ChartTextHorizontalAlignment for detail.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisdata#isbetweencategories)|Represents whether value axis crosses the category axis between categories.|
||[multiLevel](/javascript/api/excel/excel.chartaxisdata#multilevel)|Represents whether an axis is multilevel or not.|
||[numberFormat](/javascript/api/excel/excel.chartaxisdata#numberformat)|Represents the format code for the axis tick label.|
||[offset](/javascript/api/excel/excel.chartaxisdata#offset)|Represents the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.|
||[position](/javascript/api/excel/excel.chartaxisdata#position)|Represents the specified axis position where the other axis crosses. See Excel.ChartAxisPosition for details.|
||[positionAt](/javascript/api/excel/excel.chartaxisdata#positionat)|Represents the specified axis position where the other axis crosses at. You should use the SetPositionAt(double) method to set this property.|
||[textOrientation](/javascript/api/excel/excel.chartaxisdata#textorientation)|Represents the text orientation of the axis tick label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Represents chart fill formatting. Read-only.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[alignment](/javascript/api/excel/excel.chartaxisloadoptions#alignment)|Represents the alignment for the specified axis tick label. See Excel.ChartTextHorizontalAlignment for detail.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisloadoptions#isbetweencategories)|Represents whether value axis crosses the category axis between categories.|
||[multiLevel](/javascript/api/excel/excel.chartaxisloadoptions#multilevel)|Represents whether an axis is multilevel or not.|
||[numberFormat](/javascript/api/excel/excel.chartaxisloadoptions#numberformat)|Represents the format code for the axis tick label.|
||[offset](/javascript/api/excel/excel.chartaxisloadoptions#offset)|Represents the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.|
||[position](/javascript/api/excel/excel.chartaxisloadoptions#position)|Represents the specified axis position where the other axis crosses. See Excel.ChartAxisPosition for details.|
||[positionAt](/javascript/api/excel/excel.chartaxisloadoptions#positionat)|Represents the specified axis position where the other axis crosses at. You should use the SetPositionAt(double) method to set this property.|
||[textOrientation](/javascript/api/excel/excel.chartaxisloadoptions#textorientation)|Represents the text orientation of the axis tick label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|A string value that represents the formula of chart axis title using A1-style notation.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#border)|Represents the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Represents chart fill formatting.|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[border](/javascript/api/excel/excel.chartaxistitleformatdata#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[border](/javascript/api/excel/excel.chartaxistitleformatloadoptions#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[border](/javascript/api/excel/excel.chartaxistitleformatupdatedata#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[alignment](/javascript/api/excel/excel.chartaxisupdatedata#alignment)|Represents the alignment for the specified axis tick label. See Excel.ChartTextHorizontalAlignment for detail.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisupdatedata#isbetweencategories)|Represents whether value axis crosses the category axis between categories.|
||[multiLevel](/javascript/api/excel/excel.chartaxisupdatedata#multilevel)|Represents whether an axis is multilevel or not.|
||[numberFormat](/javascript/api/excel/excel.chartaxisupdatedata#numberformat)|Represents the format code for the axis tick label.|
||[offset](/javascript/api/excel/excel.chartaxisupdatedata#offset)|Represents the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.|
||[position](/javascript/api/excel/excel.chartaxisupdatedata#position)|Represents the specified axis position where the other axis crosses. See Excel.ChartAxisPosition for details.|
||[textOrientation](/javascript/api/excel/excel.chartaxisupdatedata#textorientation)|Represents the text orientation of the axis tick label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Clear the border format of a chart element.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Occurs when a chart is activated.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Occurs when a new chart is added to the worksheet.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Occurs when a chart is deactivated.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Occurs when a chart is deleted.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartcollectionloadoptions#categorylabellevel)|For EACH ITEM in the collection: Returns or sets a ChartCategoryLabelLevel enumeration constant referring to|
||[displayBlanksAs](/javascript/api/excel/excel.chartcollectionloadoptions#displayblanksas)|For EACH ITEM in the collection: Returns or sets the way that blank cells are plotted on a chart. Read/Write.|
||[plotArea](/javascript/api/excel/excel.chartcollectionloadoptions#plotarea)|For EACH ITEM in the collection: Represents the plotArea for the chart.|
||[plotBy](/javascript/api/excel/excel.chartcollectionloadoptions#plotby)|For EACH ITEM in the collection: Returns or sets the way columns or rows are used as data series on the chart. Read/Write.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartcollectionloadoptions#plotvisibleonly)|For EACH ITEM in the collection: True if only visible cells are plotted. False if both visible and hidden cells are plotted. Read/Write.|
||[seriesNameLevel](/javascript/api/excel/excel.chartcollectionloadoptions#seriesnamelevel)|For EACH ITEM in the collection: Returns or sets a ChartSeriesNameLevel enumeration constant referring to|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartcollectionloadoptions#showdatalabelsovermaximum)|For EACH ITEM in the collection: Represents whether to show the data labels when the value is greater than the maximum value on the value axis.|
||[style](/javascript/api/excel/excel.chartcollectionloadoptions#style)|For EACH ITEM in the collection: Returns or sets the chart style for the chart. Read/Write.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[categoryLabelLevel](/javascript/api/excel/excel.chartdata#categorylabellevel)|Returns or sets a ChartCategoryLabelLevel enumeration constant referring to|
||[displayBlanksAs](/javascript/api/excel/excel.chartdata#displayblanksas)|Returns or sets the way that blank cells are plotted on a chart. Read/Write.|
||[plotArea](/javascript/api/excel/excel.chartdata#plotarea)|Represents the plotArea for the chart.|
||[plotBy](/javascript/api/excel/excel.chartdata#plotby)|Returns or sets the way columns or rows are used as data series on the chart. Read/Write.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartdata#plotvisibleonly)|True if only visible cells are plotted. False if both visible and hidden cells are plotted. Read/Write.|
||[seriesNameLevel](/javascript/api/excel/excel.chartdata#seriesnamelevel)|Returns or sets a ChartSeriesNameLevel enumeration constant referring to|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartdata#showdatalabelsovermaximum)|Represents whether to show the data labels when the value is greater than the maximum value on the value axis.|
||[style](/javascript/api/excel/excel.chartdata#style)|Returns or sets the chart style for the chart. Read/Write.|
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
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[autoText](/javascript/api/excel/excel.chartdatalabeldata#autotext)|Boolean value representing if data label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.chartdatalabeldata#format)|Represents the format of chart data label.|
||[formula](/javascript/api/excel/excel.chartdatalabeldata#formula)|String value that represents the formula of chart data label using A1-style notation.|
||[height](/javascript/api/excel/excel.chartdatalabeldata#height)|Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabeldata#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.chartdatalabeldata#left)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabeldata#numberformat)|String value that represents the format code for data label.|
||[text](/javascript/api/excel/excel.chartdatalabeldata#text)|String representing the text of the data label on a chart.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabeldata#textorientation)|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.chartdatalabeldata#top)|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabeldata#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
||[width](/javascript/api/excel/excel.chartdatalabeldata#width)|Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#border)|Represents the border format, which includes color, linestyle, and weight. Read-only.|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[border](/javascript/api/excel/excel.chartdatalabelformatdata#border)|Represents the border format, which includes color, linestyle, and weight. Read-only.|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[border](/javascript/api/excel/excel.chartdatalabelformatloadoptions#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[border](/javascript/api/excel/excel.chartdatalabelformatupdatedata#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[autoText](/javascript/api/excel/excel.chartdatalabelloadoptions#autotext)|Boolean value representing if data label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.chartdatalabelloadoptions#format)|Represents the format of chart data label.|
||[formula](/javascript/api/excel/excel.chartdatalabelloadoptions#formula)|String value that represents the formula of chart data label using A1-style notation.|
||[height](/javascript/api/excel/excel.chartdatalabelloadoptions#height)|Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.chartdatalabelloadoptions#left)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#numberformat)|String value that represents the format code for data label.|
||[text](/javascript/api/excel/excel.chartdatalabelloadoptions#text)|String representing the text of the data label on a chart.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelloadoptions#textorientation)|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.chartdatalabelloadoptions#top)|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
||[width](/javascript/api/excel/excel.chartdatalabelloadoptions#width)|Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible.|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[autoText](/javascript/api/excel/excel.chartdatalabelupdatedata#autotext)|Boolean value representing if data label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.chartdatalabelupdatedata#format)|Represents the format of chart data label.|
||[formula](/javascript/api/excel/excel.chartdatalabelupdatedata#formula)|String value that represents the formula of chart data label using A1-style notation.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.chartdatalabelupdatedata#left)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#numberformat)|String value that represents the format code for data label.|
||[text](/javascript/api/excel/excel.chartdatalabelupdatedata#text)|String representing the text of the data label on a chart.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelupdatedata#textorientation)|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.chartdatalabelupdatedata#top)|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autotext)|Represents whether data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Represents the format code for data labels.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[autoText](/javascript/api/excel/excel.chartdatalabelsdata#autotext)|Represents whether data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsdata#numberformat)|Represents the format code for data labels.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsdata#textorientation)|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[autoText](/javascript/api/excel/excel.chartdatalabelsloadoptions#autotext)|Represents whether data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#numberformat)|Represents the format code for data labels.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsloadoptions#textorientation)|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[autoText](/javascript/api/excel/excel.chartdatalabelsupdatedata#autotext)|Represents whether data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#horizontalalignment)|Represents the horizontal alignment for chart data label. See Excel.ChartTextHorizontalAlignment for details.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#numberformat)|Represents the format code for data labels.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsupdatedata#textorientation)|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#verticalalignment)|Represents the vertical alignment of chart data label. See Excel.ChartTextVerticalAlignment for details.|
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
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[height](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#height)|For EACH ITEM in the collection: Represents the height of the legendEntry on the chart legend.|
||[index](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#index)|For EACH ITEM in the collection: Represents the index of the legendEntry in the chart legend.|
||[left](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#left)|For EACH ITEM in the collection: Represents the left of a chart legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#top)|For EACH ITEM in the collection: Represents the top of a chart legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#width)|For EACH ITEM in the collection: Represents the width of the legendEntry on the chart Legend.|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[height](/javascript/api/excel/excel.chartlegendentrydata#height)|Represents the height of the legendEntry on the chart legend.|
||[index](/javascript/api/excel/excel.chartlegendentrydata#index)|Represents the index of the legendEntry in the chart legend.|
||[left](/javascript/api/excel/excel.chartlegendentrydata#left)|Represents the left of a chart legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentrydata#top)|Represents the top of a chart legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentrydata#width)|Represents the width of the legendEntry on the chart Legend.|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[height](/javascript/api/excel/excel.chartlegendentryloadoptions#height)|Represents the height of the legendEntry on the chart legend.|
||[index](/javascript/api/excel/excel.chartlegendentryloadoptions#index)|Represents the index of the legendEntry in the chart legend.|
||[left](/javascript/api/excel/excel.chartlegendentryloadoptions#left)|Represents the left of a chart legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentryloadoptions#top)|Represents the top of a chart legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentryloadoptions#width)|Represents the width of the legendEntry on the chart Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#border)|Represents the border format, which includes color, linestyle, and weight. Read-only.|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[border](/javascript/api/excel/excel.chartlegendformatdata#border)|Represents the border format, which includes color, linestyle, and weight. Read-only.|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[border](/javascript/api/excel/excel.chartlegendformatloadoptions#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[border](/javascript/api/excel/excel.chartlegendformatupdatedata#border)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartloadoptions#categorylabellevel)|Returns or sets a ChartCategoryLabelLevel enumeration constant referring to|
||[displayBlanksAs](/javascript/api/excel/excel.chartloadoptions#displayblanksas)|Returns or sets the way that blank cells are plotted on a chart. Read/Write.|
||[plotArea](/javascript/api/excel/excel.chartloadoptions#plotarea)|Represents the plotArea for the chart.|
||[plotBy](/javascript/api/excel/excel.chartloadoptions#plotby)|Returns or sets the way columns or rows are used as data series on the chart. Read/Write.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartloadoptions#plotvisibleonly)|True if only visible cells are plotted. False if both visible and hidden cells are plotted. Read/Write.|
||[seriesNameLevel](/javascript/api/excel/excel.chartloadoptions#seriesnamelevel)|Returns or sets a ChartSeriesNameLevel enumeration constant referring to|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartloadoptions#showdatalabelsovermaximum)|Represents whether to show the data labels when the value is greater than the maximum value on the value axis.|
||[style](/javascript/api/excel/excel.chartloadoptions#style)|Returns or sets the chart style for the chart. Read/Write.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Represents the height value of plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Represents the insideHeight value of plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Represents the insideLeft value of plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Represents the insideTop value of plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Represents the insideWidth value of plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Represents the left value of plotArea.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Represents the position of plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Represents the formatting of a chart plotArea.|
||[set(properties: Excel.ChartPlotArea)](/javascript/api/excel/excel.chartplotarea#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartPlotAreaUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartplotarea#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Represents the top value of plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Represents the width value of plotArea.|
|[ChartPlotAreaData](/javascript/api/excel/excel.chartplotareadata)|[format](/javascript/api/excel/excel.chartplotareadata#format)|Represents the formatting of a chart plotArea.|
||[height](/javascript/api/excel/excel.chartplotareadata#height)|Represents the height value of plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotareadata#insideheight)|Represents the insideHeight value of plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotareadata#insideleft)|Represents the insideLeft value of plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotareadata#insidetop)|Represents the insideTop value of plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotareadata#insidewidth)|Represents the insideWidth value of plotArea.|
||[left](/javascript/api/excel/excel.chartplotareadata#left)|Represents the left value of plotArea.|
||[position](/javascript/api/excel/excel.chartplotareadata#position)|Represents the position of plotArea.|
||[top](/javascript/api/excel/excel.chartplotareadata#top)|Represents the top value of plotArea.|
||[width](/javascript/api/excel/excel.chartplotareadata#width)|Represents the width value of plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#border)|Represents the border attributes of a chart plotArea.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Represents the fill format of an object, which includes background formatting information.|
||[set(properties: Excel.ChartPlotAreaFormat)](/javascript/api/excel/excel.chartplotareaformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartPlotAreaFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartplotareaformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartPlotAreaFormatData](/javascript/api/excel/excel.chartplotareaformatdata)|[border](/javascript/api/excel/excel.chartplotareaformatdata#border)|Represents the border attributes of a chart plotArea.|
|[ChartPlotAreaFormatLoadOptions](/javascript/api/excel/excel.chartplotareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartplotareaformatloadoptions#$all)||
||[border](/javascript/api/excel/excel.chartplotareaformatloadoptions#border)|Represents the border attributes of a chart plotArea.|
|[ChartPlotAreaFormatUpdateData](/javascript/api/excel/excel.chartplotareaformatupdatedata)|[border](/javascript/api/excel/excel.chartplotareaformatupdatedata#border)|Represents the border attributes of a chart plotArea.|
|[ChartPlotAreaLoadOptions](/javascript/api/excel/excel.chartplotarealoadoptions)|[$all](/javascript/api/excel/excel.chartplotarealoadoptions#$all)||
||[format](/javascript/api/excel/excel.chartplotarealoadoptions#format)|Represents the formatting of a chart plotArea.|
||[height](/javascript/api/excel/excel.chartplotarealoadoptions#height)|Represents the height value of plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarealoadoptions#insideheight)|Represents the insideHeight value of plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarealoadoptions#insideleft)|Represents the insideLeft value of plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarealoadoptions#insidetop)|Represents the insideTop value of plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarealoadoptions#insidewidth)|Represents the insideWidth value of plotArea.|
||[left](/javascript/api/excel/excel.chartplotarealoadoptions#left)|Represents the left value of plotArea.|
||[position](/javascript/api/excel/excel.chartplotarealoadoptions#position)|Represents the position of plotArea.|
||[top](/javascript/api/excel/excel.chartplotarealoadoptions#top)|Represents the top value of plotArea.|
||[width](/javascript/api/excel/excel.chartplotarealoadoptions#width)|Represents the width value of plotArea.|
|[ChartPlotAreaUpdateData](/javascript/api/excel/excel.chartplotareaupdatedata)|[format](/javascript/api/excel/excel.chartplotareaupdatedata#format)|Represents the formatting of a chart plotArea.|
||[height](/javascript/api/excel/excel.chartplotareaupdatedata#height)|Represents the height value of plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotareaupdatedata#insideheight)|Represents the insideHeight value of plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotareaupdatedata#insideleft)|Represents the insideLeft value of plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotareaupdatedata#insidetop)|Represents the insideTop value of plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotareaupdatedata#insidewidth)|Represents the insideWidth value of plotArea.|
||[left](/javascript/api/excel/excel.chartplotareaupdatedata#left)|Represents the left value of plotArea.|
||[position](/javascript/api/excel/excel.chartplotareaupdatedata#position)|Represents the position of plotArea.|
||[top](/javascript/api/excel/excel.chartplotareaupdatedata#top)|Represents the top value of plotArea.|
||[width](/javascript/api/excel/excel.chartplotareaupdatedata#width)|Represents the width value of plotArea.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Returns or sets the group for the specified series. Read/Write|
||[explosion](/javascript/api/excel/excel.chartseries#explosion)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/Write.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/Write|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/Write.|
||[overlap](/javascript/api/excel/excel.chartseries#overlap)|Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts. Read/Write.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Represents a collection of all dataLabels in the series.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/Write.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/Write.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/Write.|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriescollectionloadoptions#axisgroup)|For EACH ITEM in the collection: Returns or sets the group for the specified series. Read/Write|
||[dataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#datalabels)|For EACH ITEM in the collection: Represents a collection of all dataLabels in the series.|
||[explosion](/javascript/api/excel/excel.chartseriescollectionloadoptions#explosion)|For EACH ITEM in the collection: Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/Write.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriescollectionloadoptions#firstsliceangle)|For EACH ITEM in the collection: Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/Write|
||[invertIfNegative](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertifnegative)|For EACH ITEM in the collection: True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/Write.|
||[overlap](/javascript/api/excel/excel.chartseriescollectionloadoptions#overlap)|For EACH ITEM in the collection: Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts. Read/Write.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#secondplotsize)|For EACH ITEM in the collection: Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/Write.|
||[splitType](/javascript/api/excel/excel.chartseriescollectionloadoptions#splittype)|For EACH ITEM in the collection: Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/Write.|
||[varyByCategories](/javascript/api/excel/excel.chartseriescollectionloadoptions#varybycategories)|For EACH ITEM in the collection: True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/Write.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[axisGroup](/javascript/api/excel/excel.chartseriesdata#axisgroup)|Returns or sets the group for the specified series. Read/Write|
||[dataLabels](/javascript/api/excel/excel.chartseriesdata#datalabels)|Represents a collection of all dataLabels in the series.|
||[explosion](/javascript/api/excel/excel.chartseriesdata#explosion)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/Write.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesdata#firstsliceangle)|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/Write|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesdata#invertifnegative)|True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/Write.|
||[overlap](/javascript/api/excel/excel.chartseriesdata#overlap)|Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts. Read/Write.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesdata#secondplotsize)|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/Write.|
||[splitType](/javascript/api/excel/excel.chartseriesdata#splittype)|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/Write.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesdata#varybycategories)|True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/Write.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriesloadoptions#axisgroup)|Returns or sets the group for the specified series. Read/Write|
||[dataLabels](/javascript/api/excel/excel.chartseriesloadoptions#datalabels)|Represents a collection of all dataLabels in the series.|
||[explosion](/javascript/api/excel/excel.chartseriesloadoptions#explosion)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/Write.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesloadoptions#firstsliceangle)|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/Write|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesloadoptions#invertifnegative)|True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/Write.|
||[overlap](/javascript/api/excel/excel.chartseriesloadoptions#overlap)|Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts. Read/Write.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesloadoptions#secondplotsize)|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/Write.|
||[splitType](/javascript/api/excel/excel.chartseriesloadoptions#splittype)|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/Write.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesloadoptions#varybycategories)|True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/Write.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[axisGroup](/javascript/api/excel/excel.chartseriesupdatedata#axisgroup)|Returns or sets the group for the specified series. Read/Write|
||[dataLabels](/javascript/api/excel/excel.chartseriesupdatedata#datalabels)|Represents a collection of all dataLabels in the series.|
||[explosion](/javascript/api/excel/excel.chartseriesupdatedata#explosion)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/Write.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesupdatedata#firstsliceangle)|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/Write|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesupdatedata#invertifnegative)|True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/Write.|
||[overlap](/javascript/api/excel/excel.chartseriesupdatedata#overlap)|Specifies how bars and columns are positioned. Can be a value between –100 and 100. Applies only to 2-D bar and 2-D column charts. Read/Write.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesupdatedata#secondplotsize)|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/Write.|
||[splitType](/javascript/api/excel/excel.chartseriesupdatedata#splittype)|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/Write.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesupdatedata#varybycategories)|True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/Write.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Represents the number of periods that the trendline extends forward.|
||[label](/javascript/api/excel/excel.charttrendline#label)|Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|True if the R-squared for the trendline is displayed on the chart.|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#backwardperiod)|For EACH ITEM in the collection: Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#forwardperiod)|For EACH ITEM in the collection: Represents the number of periods that the trendline extends forward.|
||[label](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#label)|For EACH ITEM in the collection: Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showequation)|For EACH ITEM in the collection: True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showrsquared)|For EACH ITEM in the collection: True if the R-squared for the trendline is displayed on the chart.|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinedata#backwardperiod)|Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinedata#forwardperiod)|Represents the number of periods that the trendline extends forward.|
||[label](/javascript/api/excel/excel.charttrendlinedata#label)|Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendlinedata#showequation)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendlinedata#showrsquared)|True if the R-squared for the trendline is displayed on the chart.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Boolean value representing if trendline label automatically generates appropriate text based on context.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|String value that represents the formula of chart trendline label using A1-style notation.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Represents the horizontal alignment for chart trendline label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|String value that represents the format code for trendline label.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Represents the format of chart trendline label.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.|
||[set(properties: Excel.ChartTrendlineLabel)](/javascript/api/excel/excel.charttrendlinelabel#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartTrendlineLabelUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.charttrendlinelabel#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|String representing the text of the trendline label on a chart.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Represents the vertical alignment of chart trendline label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[autoText](/javascript/api/excel/excel.charttrendlinelabeldata#autotext)|Boolean value representing if trendline label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.charttrendlinelabeldata#format)|Represents the format of chart trendline label.|
||[formula](/javascript/api/excel/excel.charttrendlinelabeldata#formula)|String value that represents the formula of chart trendline label using A1-style notation.|
||[height](/javascript/api/excel/excel.charttrendlinelabeldata#height)|Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#horizontalalignment)|Represents the horizontal alignment for chart trendline label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.charttrendlinelabeldata#left)|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#numberformat)|String value that represents the format code for trendline label.|
||[text](/javascript/api/excel/excel.charttrendlinelabeldata#text)|String representing the text of the trendline label on a chart.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabeldata#textorientation)|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttrendlinelabeldata#top)|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#verticalalignment)|Represents the vertical alignment of chart trendline label. See Excel.ChartTextVerticalAlignment for details.|
||[width](/javascript/api/excel/excel.charttrendlinelabeldata#width)|Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Represents the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Represents the fill format of the current chart trendline label.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Represents the font attributes (font name, font size, color, etc.) for a chart trendline label.|
||[set(properties: Excel.ChartTrendlineLabelFormat)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartTrendlineLabelFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartTrendlineLabelFormatData](/javascript/api/excel/excel.charttrendlinelabelformatdata)|[border](/javascript/api/excel/excel.charttrendlinelabelformatdata#border)|Represents the border format, which includes color, linestyle, and weight.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatdata#font)|Represents the font attributes (font name, font size, color, etc.) for a chart trendline label.|
|[ChartTrendlineLabelFormatLoadOptions](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#$all)||
||[border](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#border)|Represents the border format, which includes color, linestyle, and weight.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#font)|Represents the font attributes (font name, font size, color, etc.) for a chart trendline label.|
|[ChartTrendlineLabelFormatUpdateData](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata)|[border](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#border)|Represents the border format, which includes color, linestyle, and weight.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#font)|Represents the font attributes (font name, font size, color, etc.) for a chart trendline label.|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelloadoptions#$all)||
||[autoText](/javascript/api/excel/excel.charttrendlinelabelloadoptions#autotext)|Boolean value representing if trendline label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.charttrendlinelabelloadoptions#format)|Represents the format of chart trendline label.|
||[formula](/javascript/api/excel/excel.charttrendlinelabelloadoptions#formula)|String value that represents the formula of chart trendline label using A1-style notation.|
||[height](/javascript/api/excel/excel.charttrendlinelabelloadoptions#height)|Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#horizontalalignment)|Represents the horizontal alignment for chart trendline label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.charttrendlinelabelloadoptions#left)|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#numberformat)|String value that represents the format code for trendline label.|
||[text](/javascript/api/excel/excel.charttrendlinelabelloadoptions#text)|String representing the text of the trendline label on a chart.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelloadoptions#textorientation)|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttrendlinelabelloadoptions#top)|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#verticalalignment)|Represents the vertical alignment of chart trendline label. See Excel.ChartTextVerticalAlignment for details.|
||[width](/javascript/api/excel/excel.charttrendlinelabelloadoptions#width)|Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible.|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[autoText](/javascript/api/excel/excel.charttrendlinelabelupdatedata#autotext)|Boolean value representing if trendline label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.charttrendlinelabelupdatedata#format)|Represents the format of chart trendline label.|
||[formula](/javascript/api/excel/excel.charttrendlinelabelupdatedata#formula)|String value that represents the formula of chart trendline label using A1-style notation.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#horizontalalignment)|Represents the horizontal alignment for chart trendline label. See Excel.ChartTextHorizontalAlignment for details.|
||[left](/javascript/api/excel/excel.charttrendlinelabelupdatedata#left)|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#numberformat)|String value that represents the format code for trendline label.|
||[text](/javascript/api/excel/excel.charttrendlinelabelupdatedata#text)|String representing the text of the trendline label on a chart.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelupdatedata#textorientation)|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
||[top](/javascript/api/excel/excel.charttrendlinelabelupdatedata#top)|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#verticalalignment)|Represents the vertical alignment of chart trendline label. See Excel.ChartTextVerticalAlignment for details.|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#backwardperiod)|Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#forwardperiod)|Represents the number of periods that the trendline extends forward.|
||[label](/javascript/api/excel/excel.charttrendlineloadoptions#label)|Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendlineloadoptions#showequation)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendlineloadoptions#showrsquared)|True if the R-squared for the trendline is displayed on the chart.|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#backwardperiod)|Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#forwardperiod)|Represents the number of periods that the trendline extends forward.|
||[label](/javascript/api/excel/excel.charttrendlineupdatedata#label)|Represents the label of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendlineupdatedata#showequation)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendlineupdatedata#showrsquared)|True if the R-squared for the trendline is displayed on the chart.|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[categoryLabelLevel](/javascript/api/excel/excel.chartupdatedata#categorylabellevel)|Returns or sets a ChartCategoryLabelLevel enumeration constant referring to|
||[displayBlanksAs](/javascript/api/excel/excel.chartupdatedata#displayblanksas)|Returns or sets the way that blank cells are plotted on a chart. Read/Write.|
||[plotArea](/javascript/api/excel/excel.chartupdatedata#plotarea)|Represents the plotArea for the chart.|
||[plotBy](/javascript/api/excel/excel.chartupdatedata#plotby)|Returns or sets the way columns or rows are used as data series on the chart. Read/Write.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartupdatedata#plotvisibleonly)|True if only visible cells are plotted. False if both visible and hidden cells are plotted. Read/Write.|
||[seriesNameLevel](/javascript/api/excel/excel.chartupdatedata#seriesnamelevel)|Returns or sets a ChartSeriesNameLevel enumeration constant referring to|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartupdatedata#showdatalabelsovermaximum)|Represents whether to show the data labels when the value is greater than the maximum value on the value axis.|
||[style](/javascript/api/excel/excel.chartupdatedata#style)|Returns or sets the chart style for the chart. Read/Write.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|A custom data validation formula. This creates special input rules, such as preventing duplicates, or limiting the total in a range of cells.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Position of the DataPivotHierarchy.|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|Id of the DataPivotHierarchy.|
||[set(properties: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchy#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.DataPivotHierarchyUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.datapivothierarchy#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Reset the DataPivotHierarchy back to its default values.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Determines whether the data should be shown as a specific summary calculation or not.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Determines whether to show all items of the DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Gets a DataPivotHierarchy by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Gets a DataPivotHierarchy by name. If the DataPivotHierarchy does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Removes the PivotHierarchy from the current axis.|
|[DataPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#$all)||
||[field](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#field)|For EACH ITEM in the collection: Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#id)|For EACH ITEM in the collection: Id of the DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#name)|For EACH ITEM in the collection: Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#numberformat)|For EACH ITEM in the collection: Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#position)|For EACH ITEM in the collection: Position of the DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#showas)|For EACH ITEM in the collection: Determines whether the data should be shown as a specific summary calculation or not.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#summarizeby)|For EACH ITEM in the collection: Determines whether to show all items of the DataPivotHierarchy.|
|[DataPivotHierarchyData](/javascript/api/excel/excel.datapivothierarchydata)|[field](/javascript/api/excel/excel.datapivothierarchydata#field)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchydata#id)|Id of the DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchydata#name)|Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchydata#numberformat)|Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchydata#position)|Position of the DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchydata#showas)|Determines whether the data should be shown as a specific summary calculation or not.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchydata#summarizeby)|Determines whether to show all items of the DataPivotHierarchy.|
|[DataPivotHierarchyLoadOptions](/javascript/api/excel/excel.datapivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchyloadoptions#$all)||
||[field](/javascript/api/excel/excel.datapivothierarchyloadoptions#field)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchyloadoptions#id)|Id of the DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchyloadoptions#name)|Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyloadoptions#numberformat)|Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchyloadoptions#position)|Position of the DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchyloadoptions#showas)|Determines whether the data should be shown as a specific summary calculation or not.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyloadoptions#summarizeby)|Determines whether to show all items of the DataPivotHierarchy.|
|[DataPivotHierarchyUpdateData](/javascript/api/excel/excel.datapivothierarchyupdatedata)|[field](/javascript/api/excel/excel.datapivothierarchyupdatedata#field)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchyupdatedata#name)|Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyupdatedata#numberformat)|Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchyupdatedata#position)|Position of the DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchyupdatedata#showas)|Determines whether the data should be shown as a specific summary calculation or not.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyupdatedata#summarizeby)|Determines whether to show all items of the DataPivotHierarchy.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Clears the data validation from the current range.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Error alert when user enters invalid data.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Prompt when users select a cell.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Type of the data validation, see Excel.DataValidationType for details.|
||[valid](/javascript/api/excel/excel.datavalidation#valid)|Represents if all cell values are valid according to the data validation rules.|
||[rule](/javascript/api/excel/excel.datavalidation#rule)|Data validation rule that contains different type of data validation criteria.|
||[set(properties: Excel.DataValidation)](/javascript/api/excel/excel.datavalidation#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.DataValidationUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.datavalidation#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[DataValidationData](/javascript/api/excel/excel.datavalidationdata)|[errorAlert](/javascript/api/excel/excel.datavalidationdata#erroralert)|Error alert when user enters invalid data.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationdata#ignoreblanks)|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|
||[prompt](/javascript/api/excel/excel.datavalidationdata#prompt)|Prompt when users select a cell.|
||[rule](/javascript/api/excel/excel.datavalidationdata#rule)|Data validation rule that contains different type of data validation criteria.|
||[type](/javascript/api/excel/excel.datavalidationdata#type)|Type of the data validation, see Excel.DataValidationType for details.|
||[valid](/javascript/api/excel/excel.datavalidationdata#valid)|Represents if all cell values are valid according to the data validation rules.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Represents error alert message.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Determines whether to show an error alert dialog or not when a user enters invalid data. The default is true.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Represents data validation alert type, please see Excel.DataValidationAlertStyle for details.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Represents error alert dialog title.|
|[DataValidationLoadOptions](/javascript/api/excel/excel.datavalidationloadoptions)|[$all](/javascript/api/excel/excel.datavalidationloadoptions#$all)||
||[errorAlert](/javascript/api/excel/excel.datavalidationloadoptions#erroralert)|Error alert when user enters invalid data.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationloadoptions#ignoreblanks)|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|
||[prompt](/javascript/api/excel/excel.datavalidationloadoptions#prompt)|Prompt when users select a cell.|
||[rule](/javascript/api/excel/excel.datavalidationloadoptions#rule)|Data validation rule that contains different type of data validation criteria.|
||[type](/javascript/api/excel/excel.datavalidationloadoptions#type)|Type of the data validation, see Excel.DataValidationType for details.|
||[valid](/javascript/api/excel/excel.datavalidationloadoptions#valid)|Represents if all cell values are valid according to the data validation rules.|
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
|[DataValidationUpdateData](/javascript/api/excel/excel.datavalidationupdatedata)|[errorAlert](/javascript/api/excel/excel.datavalidationupdatedata#erroralert)|Error alert when user enters invalid data.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationupdatedata#ignoreblanks)|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|
||[prompt](/javascript/api/excel/excel.datavalidationupdatedata#prompt)|Prompt when users select a cell.|
||[rule](/javascript/api/excel/excel.datavalidationupdatedata#rule)|Data validation rule that contains different type of data validation criteria.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell). With the ternary operators Between and NotBetween, specifies the lower bound operand.|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|With the ternary operators Between and NotBetween, specifies the upper bound operand. Is not used with the binary operators, such as GreaterThan.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|The operator to use for validating the data.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Determines whether to allow multiple filter items.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Position of the FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Returns the PivotFields associated with the FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|Id of the FilterPivotHierarchy.|
||[set(properties: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchy#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.FilterPivotHierarchyUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.filterpivothierarchy#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Reset the FilterPivotHierarchy back to its default values.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Gets a FilterPivotHierarchy by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Gets a FilterPivotHierarchy by name. If the FilterPivotHierarchy does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Removes the PivotHierarchy from the current axis.|
|[FilterPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#enablemultiplefilteritems)|For EACH ITEM in the collection: Determines whether to allow multiple filter items.|
||[id](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#id)|For EACH ITEM in the collection: Id of the FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#name)|For EACH ITEM in the collection: Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#position)|For EACH ITEM in the collection: Position of the FilterPivotHierarchy.|
|[FilterPivotHierarchyData](/javascript/api/excel/excel.filterpivothierarchydata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchydata#enablemultiplefilteritems)|Determines whether to allow multiple filter items.|
||[fields](/javascript/api/excel/excel.filterpivothierarchydata#fields)|Returns the PivotFields associated with the FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchydata#id)|Id of the FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchydata#name)|Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchydata#position)|Position of the FilterPivotHierarchy.|
|[FilterPivotHierarchyLoadOptions](/javascript/api/excel/excel.filterpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchyloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyloadoptions#enablemultiplefilteritems)|Determines whether to allow multiple filter items.|
||[id](/javascript/api/excel/excel.filterpivothierarchyloadoptions#id)|Id of the FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchyloadoptions#name)|Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchyloadoptions#position)|Position of the FilterPivotHierarchy.|
|[FilterPivotHierarchyUpdateData](/javascript/api/excel/excel.filterpivothierarchyupdatedata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyupdatedata#enablemultiplefilteritems)|Determines whether to allow multiple filter items.|
||[name](/javascript/api/excel/excel.filterpivothierarchyupdatedata#name)|Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchyupdatedata#position)|Position of the FilterPivotHierarchy.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Displays the list in cell drop down or not, it defaults to true.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Source of the list for data validation|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Name of the PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|Id of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Returns the PivotItems that comprise with the PivotField.|
||[set(properties: Excel.PivotField)](/javascript/api/excel/excel.pivotfield#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.PivotFieldUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.pivotfield#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Determines whether to show all items of the PivotField.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Subtotals of the PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Gets the number of pivot fields in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Gets a PivotField by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Gets a PivotField by name. If the PivotField does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Gets the loaded child items in this collection.|
|[PivotFieldCollectionLoadOptions](/javascript/api/excel/excel.pivotfieldcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#id)|For EACH ITEM in the collection: Id of the PivotField.|
||[name](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#name)|For EACH ITEM in the collection: Name of the PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#showallitems)|For EACH ITEM in the collection: Determines whether to show all items of the PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#subtotals)|For EACH ITEM in the collection: Subtotals of the PivotField.|
|[PivotFieldData](/javascript/api/excel/excel.pivotfielddata)|[id](/javascript/api/excel/excel.pivotfielddata#id)|Id of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfielddata#items)|Returns the PivotFields associated with the PivotField.|
||[name](/javascript/api/excel/excel.pivotfielddata#name)|Name of the PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfielddata#showallitems)|Determines whether to show all items of the PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfielddata#subtotals)|Subtotals of the PivotField.|
|[PivotFieldLoadOptions](/javascript/api/excel/excel.pivotfieldloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldloadoptions#id)|Id of the PivotField.|
||[name](/javascript/api/excel/excel.pivotfieldloadoptions#name)|Name of the PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldloadoptions#showallitems)|Determines whether to show all items of the PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfieldloadoptions#subtotals)|Subtotals of the PivotField.|
|[PivotFieldUpdateData](/javascript/api/excel/excel.pivotfieldupdatedata)|[name](/javascript/api/excel/excel.pivotfieldupdatedata#name)|Name of the PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldupdatedata#showallitems)|Determines whether to show all items of the PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfieldupdatedata#subtotals)|Subtotals of the PivotField.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Name of the PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Returns the PivotFields associated with the PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|Id of the PivotHierarchy.|
||[set(properties: Excel.PivotHierarchy)](/javascript/api/excel/excel.pivothierarchy#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.PivotHierarchyUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.pivothierarchy#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Gets a PivotHierarchy by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Gets the loaded child items in this collection.|
|[PivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.pivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#id)|For EACH ITEM in the collection: Id of the PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#name)|For EACH ITEM in the collection: Name of the PivotHierarchy.|
|[PivotHierarchyData](/javascript/api/excel/excel.pivothierarchydata)|[fields](/javascript/api/excel/excel.pivothierarchydata#fields)|Returns the PivotFields associated with the PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchydata#id)|Id of the PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchydata#name)|Name of the PivotHierarchy.|
|[PivotHierarchyLoadOptions](/javascript/api/excel/excel.pivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchyloadoptions#id)|Id of the PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchyloadoptions#name)|Name of the PivotHierarchy.|
|[PivotHierarchyUpdateData](/javascript/api/excel/excel.pivothierarchyupdatedata)|[name](/javascript/api/excel/excel.pivothierarchyupdatedata#name)|Name of the PivotHierarchy.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Name of the PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|Id of the PivotItem.|
||[set(properties: Excel.PivotItem)](/javascript/api/excel/excel.pivotitem#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.PivotItemUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.pivotitem#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Determines whether the PivotItem is visible or not.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Gets the number of pivot items in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Gets a PivotItem by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Gets a PivotItem by name. If the PivotItem does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Gets the loaded child items in this collection.|
|[PivotItemCollectionLoadOptions](/javascript/api/excel/excel.pivotitemcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotitemcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemcollectionloadoptions#id)|For EACH ITEM in the collection: Id of the PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitemcollectionloadoptions#isexpanded)|For EACH ITEM in the collection: Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitemcollectionloadoptions#name)|For EACH ITEM in the collection: Name of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemcollectionloadoptions#visible)|For EACH ITEM in the collection: Determines whether the PivotItem is visible or not.|
|[PivotItemData](/javascript/api/excel/excel.pivotitemdata)|[id](/javascript/api/excel/excel.pivotitemdata#id)|Id of the PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitemdata#isexpanded)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitemdata#name)|Name of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemdata#visible)|Determines whether the PivotItem is visible or not.|
|[PivotItemLoadOptions](/javascript/api/excel/excel.pivotitemloadoptions)|[$all](/javascript/api/excel/excel.pivotitemloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemloadoptions#id)|Id of the PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitemloadoptions#isexpanded)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitemloadoptions#name)|Name of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemloadoptions#visible)|Determines whether the PivotItem is visible or not.|
|[PivotItemUpdateData](/javascript/api/excel/excel.pivotitemupdatedata)|[isExpanded](/javascript/api/excel/excel.pivotitemupdatedata#isexpanded)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitemupdatedata#name)|Name of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemupdatedata#visible)|Determines whether the PivotItem is visible or not.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Returns the range where the PivotTable's column labels reside.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Returns the range where the PivotTable's data values reside.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Returns the range of the PivotTable's filter area.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Returns the range the PivotTable exists on, excluding the filter area.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Returns the range where the PivotTable's row labels reside.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|
||[set(properties: Excel.PivotLayout)](/javascript/api/excel/excel.pivotlayout#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.PivotLayoutUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.pivotlayout#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Specifies whether the PivotTable report shows grand totals for columns.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Specifies whether the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null.|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[layoutType](/javascript/api/excel/excel.pivotlayoutdata#layouttype)|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showcolumngrandtotals)|Specifies whether the PivotTable report shows grand totals for columns.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showrowgrandtotals)|Specifies whether the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutdata#subtotallocation)|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null.|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[$all](/javascript/api/excel/excel.pivotlayoutloadoptions#$all)||
||[layoutType](/javascript/api/excel/excel.pivotlayoutloadoptions#layouttype)|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showcolumngrandtotals)|Specifies whether the PivotTable report shows grand totals for columns.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showrowgrandtotals)|Specifies whether the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutloadoptions#subtotallocation)|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null.|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[layoutType](/javascript/api/excel/excel.pivotlayoutupdatedata#layouttype)|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showcolumngrandtotals)|Specifies whether the PivotTable report shows grand totals for columns.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showrowgrandtotals)|Specifies whether the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutupdatedata#subtotallocation)|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Deletes the PivotTable.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|The Column Pivot Hierarchies of the PivotTable.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|The Data Pivot Hierarchies of the PivotTable.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|The Filter Pivot Hierarchies of the PivotTable.|
||[hierarchies](/javascript/api/excel/excel.pivottable#hierarchies)|The Pivot Hierarchies of the PivotTable.|
||[layout](/javascript/api/excel/excel.pivottable#layout)|The PivotLayout describing the layout and visual structure of the PivotTable.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|The Row Pivot Hierarchies of the PivotTable.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Add a Pivottable based on the specified source data and insert it at the top left cell of the destination range.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[layout](/javascript/api/excel/excel.pivottablecollectionloadoptions#layout)|For EACH ITEM in the collection: The PivotLayout describing the layout and visual structure of the PivotTable.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[columnHierarchies](/javascript/api/excel/excel.pivottabledata#columnhierarchies)|The Column Pivot Hierarchies of the PivotTable.|
||[dataHierarchies](/javascript/api/excel/excel.pivottabledata#datahierarchies)|The Data Pivot Hierarchies of the PivotTable.|
||[filterHierarchies](/javascript/api/excel/excel.pivottabledata#filterhierarchies)|The Filter Pivot Hierarchies of the PivotTable.|
||[hierarchies](/javascript/api/excel/excel.pivottabledata#hierarchies)|The Pivot Hierarchies of the PivotTable.|
||[rowHierarchies](/javascript/api/excel/excel.pivottabledata#rowhierarchies)|The Row Pivot Hierarchies of the PivotTable.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[layout](/javascript/api/excel/excel.pivottableloadoptions#layout)|The PivotLayout describing the layout and visual structure of the PivotTable.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Returns a data validation object.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[dataValidation](/javascript/api/excel/excel.rangedata#datavalidation)|Returns a data validation object.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[dataValidation](/javascript/api/excel/excel.rangeloadoptions#datavalidation)|Returns a data validation object.|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeupdatedata#datavalidation)|Returns a data validation object.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Position of the RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Returns the PivotFields associated with the RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|Id of the RowColumnPivotHierarchy.|
||[set(properties: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RowColumnPivotHierarchyUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Reset the RowColumnPivotHierarchy back to its default values.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Gets a RowColumnPivotHierarchy by its name or id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Gets the loaded child items in this collection.|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Removes the PivotHierarchy from the current axis.|
|[RowColumnPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#id)|For EACH ITEM in the collection: Id of the RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#name)|For EACH ITEM in the collection: Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#position)|For EACH ITEM in the collection: Position of the RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyData](/javascript/api/excel/excel.rowcolumnpivothierarchydata)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchydata#fields)|Returns the PivotFields associated with the RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchydata#id)|Id of the RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchydata#name)|Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchydata#position)|Position of the RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#id)|Id of the RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#name)|Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#position)|Position of the RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyUpdateData](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#name)|Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#position)|Position of the RowColumnPivotHierarchy.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Toggle JavaScript events in the current task pane or content add-in.|
|[RuntimeData](/javascript/api/excel/excel.runtimedata)|[enableEvents](/javascript/api/excel/excel.runtimedata#enableevents)|Toggle JavaScript events in the current task pane or content add-in.|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[enableEvents](/javascript/api/excel/excel.runtimeloadoptions#enableevents)|Toggle JavaScript events in the current task pane or content add-in.|
|[RuntimeUpdateData](/javascript/api/excel/excel.runtimeupdatedata)|[enableEvents](/javascript/api/excel/excel.runtimeupdatedata#enableevents)|Toggle JavaScript events in the current task pane or content add-in.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|The base PivotField to base the ShowAs calculation, if applicable based on the ShowAsCalculation type, else null.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|The base Item to base the ShowAs calculation on, if applicable based on the ShowAsCalculation type, else null.|
||[calculation](/javascript/api/excel/excel.showasrule#calculation)|The ShowAs Calculation to use for the Data PivotField. See Excel.ShowAsCalculation for Details.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|The text orientation for the style.|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[autoIndent](/javascript/api/excel/excel.stylecollectionloadoptions#autoindent)|For EACH ITEM in the collection: Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/javascript/api/excel/excel.stylecollectionloadoptions#textorientation)|For EACH ITEM in the collection: The text orientation for the style.|
|[StyleData](/javascript/api/excel/excel.styledata)|[autoIndent](/javascript/api/excel/excel.styledata#autoindent)|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/javascript/api/excel/excel.styledata#textorientation)|The text orientation for the style.|
|[StyleLoadOptions](/javascript/api/excel/excel.styleloadoptions)|[autoIndent](/javascript/api/excel/excel.styleloadoptions#autoindent)|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/javascript/api/excel/excel.styleloadoptions#textorientation)|The text orientation for the style.|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[autoIndent](/javascript/api/excel/excel.styleupdatedata#autoindent)|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/javascript/api/excel/excel.styleupdatedata#textorientation)|The text orientation for the style.|
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
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[legacyId](/javascript/api/excel/excel.tablecollectionloadoptions#legacyid)|For EACH ITEM in the collection: Returns a numeric id.|
|[TableData](/javascript/api/excel/excel.tabledata)|[legacyId](/javascript/api/excel/excel.tabledata#legacyid)|Returns a numeric id.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[legacyId](/javascript/api/excel/excel.tableloadoptions#legacyid)|Returns a numeric id.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|True if the workbook is open in Read-only mode. Read-only.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[WorkbookData](/javascript/api/excel/excel.workbookdata)|[readOnly](/javascript/api/excel/excel.workbookdata#readonly)|True if the workbook is open in Read-only mode. Read-only.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[readOnly](/javascript/api/excel/excel.workbookloadoptions#readonly)|True if the workbook is open in Read-only mode. Read-only.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|Occurs when the worksheet is calculated.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|Gets or sets the worksheet's gridlines flag.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|Gets or sets the worksheet's headings flag.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Gets the id of the worksheet that is calculated.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Gets the range that represents the changed area of a specific worksheet. It might return null object.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Occurs when any worksheet in the workbook is calculated.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[showGridlines](/javascript/api/excel/excel.worksheetcollectionloadoptions#showgridlines)|For EACH ITEM in the collection: Gets or sets the worksheet's gridlines flag.|
||[showHeadings](/javascript/api/excel/excel.worksheetcollectionloadoptions#showheadings)|For EACH ITEM in the collection: Gets or sets the worksheet's headings flag.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[showGridlines](/javascript/api/excel/excel.worksheetdata#showgridlines)|Gets or sets the worksheet's gridlines flag.|
||[showHeadings](/javascript/api/excel/excel.worksheetdata#showheadings)|Gets or sets the worksheet's headings flag.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[showGridlines](/javascript/api/excel/excel.worksheetloadoptions#showgridlines)|Gets or sets the worksheet's gridlines flag.|
||[showHeadings](/javascript/api/excel/excel.worksheetloadoptions#showheadings)|Gets or sets the worksheet's headings flag.|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[showGridlines](/javascript/api/excel/excel.worksheetupdatedata#showgridlines)|Gets or sets the worksheet's gridlines flag.|
||[showHeadings](/javascript/api/excel/excel.worksheetupdatedata#showheadings)|Gets or sets the worksheet's headings flag.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel&view=excel-js-1.8)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
