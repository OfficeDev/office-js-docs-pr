---
title: Excel JavaScript API requirement set 1.1
description: 'Details about the ExcelApi 1.1 requirement set'
ms.date: 07/25/2019
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript API requirement set 1.1

Excel JavaScript API 1.1 is the first version of the API. It is the only Excel-specific requirement set supported by Excel 2016.

## API list

To see a complete list of all APIs supported by this requirement set (including previously released APIs), [click here to see a version-specific of the API reference documentation]((/javascript/api/excel?view=excel-js-1.1)).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: "Recalculate" \| "Full" \| "FullRebuild")](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalculate all currently opened workbooks in Excel.|
||[calculate(calculationType: Excel.CalculationType)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalculate all currently opened workbooks in Excel.|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|Returns the calculation mode used in the workbook, as defined by the constants in Excel.CalculationMode. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.|
||[set(properties: Excel.Application)](/javascript/api/excel/excel.application#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ApplicationUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.application#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ApplicationData](/javascript/api/excel/excel.applicationdata)|[calculationMode](/javascript/api/excel/excel.applicationdata#calculationmode)|Returns the calculation mode used in the workbook, as defined by the constants in Excel.CalculationMode. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.|
|[ApplicationLoadOptions](/javascript/api/excel/excel.applicationloadoptions)|[$all](/javascript/api/excel/excel.applicationloadoptions#$all)||
||[calculationMode](/javascript/api/excel/excel.applicationloadoptions#calculationmode)|Returns the calculation mode used in the workbook, as defined by the constants in Excel.CalculationMode. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.|
|[ApplicationUpdateData](/javascript/api/excel/excel.applicationupdatedata)|[calculationMode](/javascript/api/excel/excel.applicationupdatedata#calculationmode)|Returns the calculation mode used in the workbook, as defined by the constants in Excel.CalculationMode. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Returns the range represented by the binding. Will throw an error if binding is not of the correct type.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Returns the table represented by the binding. Will throw an error if binding is not of the correct type.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Returns the text represented by the binding. Will throw an error if binding is not of the correct type.|
||[id](/javascript/api/excel/excel.binding#id)|Represents binding identifier. Read-only.|
||[type](/javascript/api/excel/excel.binding#type)|Returns the type of the binding. See Excel.BindingType for details. Read-only.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Gets a binding object by ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Gets a binding object based on its position in the items array.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Returns the number of bindings in the collection. Read-only.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Gets the loaded child items in this collection.|
|[BindingCollectionLoadOptions](/javascript/api/excel/excel.bindingcollectionloadoptions)|[$all](/javascript/api/excel/excel.bindingcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingcollectionloadoptions#id)|For EACH ITEM in the collection: Represents binding identifier. Read-only.|
||[type](/javascript/api/excel/excel.bindingcollectionloadoptions#type)|For EACH ITEM in the collection: Returns the type of the binding. See Excel.BindingType for details. Read-only.|
|[BindingData](/javascript/api/excel/excel.bindingdata)|[id](/javascript/api/excel/excel.bindingdata#id)|Represents binding identifier. Read-only.|
||[type](/javascript/api/excel/excel.bindingdata#type)|Returns the type of the binding. See Excel.BindingType for details. Read-only.|
|[BindingLoadOptions](/javascript/api/excel/excel.bindingloadoptions)|[$all](/javascript/api/excel/excel.bindingloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingloadoptions#id)|Represents binding identifier. Read-only.|
||[type](/javascript/api/excel/excel.bindingloadoptions#type)|Returns the type of the binding. See Excel.BindingType for details. Read-only.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Deletes the chart object.|
||[height](/javascript/api/excel/excel.chart#height)|Represents the height, in points, of the chart object.|
||[left](/javascript/api/excel/excel.chart#left)|The distance, in points, from the left side of the chart to the worksheet origin.|
||[name](/javascript/api/excel/excel.chart#name)|Represents the name of a chart object.|
||[axes](/javascript/api/excel/excel.chart#axes)|Represents chart axes. Read-only.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Represents the datalabels on the chart. Read-only.|
||[format](/javascript/api/excel/excel.chart#format)|Encapsulates the format properties for the chart area. Read-only.|
||[legend](/javascript/api/excel/excel.chart#legend)|Represents the legend for the chart. Read-only.|
||[series](/javascript/api/excel/excel.chart#series)|Represents either a single series or collection of series in the chart. Read-only.|
||[title](/javascript/api/excel/excel.chart#title)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title. Read-only.|
||[set(properties: Excel.Chart)](/javascript/api/excel/excel.chart#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chart#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[setData(sourceData: Range, seriesBy?: "Auto" \| "Columns" \| "Rows")](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Resets the source data for the chart.|
||[setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Resets the source data for the chart.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Positions the chart relative to cells on the worksheet.|
||[top](/javascript/api/excel/excel.chart#top)|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/javascript/api/excel/excel.chart#width)|Represents the width, in points, of the chart object.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Represents the fill format of an object, which includes background formatting information. Read-only.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.|
||[set(properties: Excel.ChartAreaFormat)](/javascript/api/excel/excel.chartareaformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartAreaFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartareaformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[font](/javascript/api/excel/excel.chartareaformatdata#font)|Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartareaformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartareaformatloadoptions#font)|Represents the font attributes (font name, font size, color, etc.) for the current object.|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[font](/javascript/api/excel/excel.chartareaformatupdatedata#font)|Represents the font attributes (font name, font size, color, etc.) for the current object.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|Represents the category axis in a chart. Read-only.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|Represents the series axis of a 3-dimensional chart. Read-only.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Represents the value axis in an axis. Read-only.|
||[set(properties: Excel.ChartAxes)](/javascript/api/excel/excel.chartaxes#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartAxesUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartaxes#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartAxesData](/javascript/api/excel/excel.chartaxesdata)|[categoryAxis](/javascript/api/excel/excel.chartaxesdata#categoryaxis)|Represents the category axis in a chart. Read-only.|
||[seriesAxis](/javascript/api/excel/excel.chartaxesdata#seriesaxis)|Represents the series axis of a 3-dimensional chart. Read-only.|
||[valueAxis](/javascript/api/excel/excel.chartaxesdata#valueaxis)|Represents the value axis in an axis. Read-only.|
|[ChartAxesLoadOptions](/javascript/api/excel/excel.chartaxesloadoptions)|[$all](/javascript/api/excel/excel.chartaxesloadoptions#$all)||
||[categoryAxis](/javascript/api/excel/excel.chartaxesloadoptions#categoryaxis)|Represents the category axis in a chart.|
||[seriesAxis](/javascript/api/excel/excel.chartaxesloadoptions#seriesaxis)|Represents the series axis of a 3-dimensional chart.|
||[valueAxis](/javascript/api/excel/excel.chartaxesloadoptions#valueaxis)|Represents the value axis in an axis.|
|[ChartAxesUpdateData](/javascript/api/excel/excel.chartaxesupdatedata)|[categoryAxis](/javascript/api/excel/excel.chartaxesupdatedata#categoryaxis)|Represents the category axis in a chart.|
||[seriesAxis](/javascript/api/excel/excel.chartaxesupdatedata#seriesaxis)|Represents the series axis of a 3-dimensional chart.|
||[valueAxis](/javascript/api/excel/excel.chartaxesupdatedata#valueaxis)|Represents the value axis in an axis.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Represents the formatting of a chart object, which includes line and font formatting. Read-only.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Returns a Gridlines object that represents the major gridlines for the specified axis. Read-only.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Represents the axis title. Read-only.|
||[set(properties: Excel.ChartAxis)](/javascript/api/excel/excel.chartaxis#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartAxisUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartaxis#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[format](/javascript/api/excel/excel.chartaxisdata#format)|Represents the formatting of a chart object, which includes line and font formatting. Read-only.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisdata#majorgridlines)|Returns a Gridlines object that represents the major gridlines for the specified axis. Read-only.|
||[majorUnit](/javascript/api/excel/excel.chartaxisdata#majorunit)|Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.|
||[maximum](/javascript/api/excel/excel.chartaxisdata#maximum)|Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minimum](/javascript/api/excel/excel.chartaxisdata#minimum)|Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisdata#minorgridlines)|Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.|
||[minorUnit](/javascript/api/excel/excel.chartaxisdata#minorunit)|Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.|
||[title](/javascript/api/excel/excel.chartaxisdata#title)|Represents the axis title. Read-only.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Represents chart line formatting. Read-only.|
||[set(properties: Excel.ChartAxisFormat)](/javascript/api/excel/excel.chartaxisformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartAxisFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartaxisformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartAxisFormatData](/javascript/api/excel/excel.chartaxisformatdata)|[font](/javascript/api/excel/excel.chartaxisformatdata#font)|Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.|
||[line](/javascript/api/excel/excel.chartaxisformatdata#line)|Represents chart line formatting. Read-only.|
|[ChartAxisFormatLoadOptions](/javascript/api/excel/excel.chartaxisformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxisformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartaxisformatloadoptions#font)|Represents the font attributes (font name, font size, color, etc.) for a chart axis element.|
||[line](/javascript/api/excel/excel.chartaxisformatloadoptions#line)|Represents chart line formatting.|
|[ChartAxisFormatUpdateData](/javascript/api/excel/excel.chartaxisformatupdatedata)|[font](/javascript/api/excel/excel.chartaxisformatupdatedata#font)|Represents the font attributes (font name, font size, color, etc.) for a chart axis element.|
||[line](/javascript/api/excel/excel.chartaxisformatupdatedata#line)|Represents chart line formatting.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[$all](/javascript/api/excel/excel.chartaxisloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxisloadoptions#format)|Represents the formatting of a chart object, which includes line and font formatting.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#majorgridlines)|Returns a Gridlines object that represents the major gridlines for the specified axis.|
||[majorUnit](/javascript/api/excel/excel.chartaxisloadoptions#majorunit)|Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.|
||[maximum](/javascript/api/excel/excel.chartaxisloadoptions#maximum)|Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minimum](/javascript/api/excel/excel.chartaxisloadoptions#minimum)|Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#minorgridlines)|Returns a Gridlines object that represents the minor gridlines for the specified axis.|
||[minorUnit](/javascript/api/excel/excel.chartaxisloadoptions#minorunit)|Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.|
||[title](/javascript/api/excel/excel.chartaxisloadoptions#title)|Represents the axis title.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Represents the formatting of chart axis title. Read-only.|
||[set(properties: Excel.ChartAxisTitle)](/javascript/api/excel/excel.chartaxistitle#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartAxisTitleUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartaxistitle#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Represents the axis title.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|A boolean that specifies the visibility of an axis title.|
|[ChartAxisTitleData](/javascript/api/excel/excel.chartaxistitledata)|[format](/javascript/api/excel/excel.chartaxistitledata#format)|Represents the formatting of chart axis title. Read-only.|
||[text](/javascript/api/excel/excel.chartaxistitledata#text)|Represents the axis title.|
||[visible](/javascript/api/excel/excel.chartaxistitledata#visible)|A boolean that specifies the visibility of an axis title.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.|
||[set(properties: Excel.ChartAxisTitleFormat)](/javascript/api/excel/excel.chartaxistitleformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartAxisTitleFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartaxistitleformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[font](/javascript/api/excel/excel.chartaxistitleformatdata#font)|Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartaxistitleformatloadoptions#font)|Represents the font attributes, such as font name, font size, color, etc. of chart axis title object.|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[font](/javascript/api/excel/excel.chartaxistitleformatupdatedata#font)|Represents the font attributes, such as font name, font size, color, etc. of chart axis title object.|
|[ChartAxisTitleLoadOptions](/javascript/api/excel/excel.chartaxistitleloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxistitleloadoptions#format)|Represents the formatting of chart axis title.|
||[text](/javascript/api/excel/excel.chartaxistitleloadoptions#text)|Represents the axis title.|
||[visible](/javascript/api/excel/excel.chartaxistitleloadoptions#visible)|A boolean that specifies the visibility of an axis title.|
|[ChartAxisTitleUpdateData](/javascript/api/excel/excel.chartaxistitleupdatedata)|[format](/javascript/api/excel/excel.chartaxistitleupdatedata#format)|Represents the formatting of chart axis title.|
||[text](/javascript/api/excel/excel.chartaxistitleupdatedata#text)|Represents the axis title.|
||[visible](/javascript/api/excel/excel.chartaxistitleupdatedata#visible)|A boolean that specifies the visibility of an axis title.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[format](/javascript/api/excel/excel.chartaxisupdatedata#format)|Represents the formatting of a chart object, which includes line and font formatting.|
||[majorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#majorgridlines)|Returns a Gridlines object that represents the major gridlines for the specified axis.|
||[majorUnit](/javascript/api/excel/excel.chartaxisupdatedata#majorunit)|Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.|
||[maximum](/javascript/api/excel/excel.chartaxisupdatedata#maximum)|Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minimum](/javascript/api/excel/excel.chartaxisupdatedata#minimum)|Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
||[minorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#minorgridlines)|Returns a Gridlines object that represents the minor gridlines for the specified axis.|
||[minorUnit](/javascript/api/excel/excel.chartaxisupdatedata#minorunit)|Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.|
||[title](/javascript/api/excel/excel.chartaxisupdatedata#title)|Represents the axis title.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: "Invalid" \| "ColumnClustered" \| "ColumnStacked" \| "ColumnStacked100" \| "3DColumnClustered" \| "3DColumnStacked" \| "3DColumnStacked100" \| "BarClustered" \| "BarStacked" \| "BarStacked100" \| "3DBarClustered" \| "3DBarStacked" \| "3DBarStacked100" \| "LineStacked" \| "LineStacked100" \| "LineMarkers" \| "LineMarkersStacked" \| "LineMarkersStacked100" \| "PieOfPie" \| "PieExploded" \| "3DPieExploded" \| "BarOfPie" \| "XYScatterSmooth" \| "XYScatterSmoothNoMarkers" \| "XYScatterLines" \| "XYScatterLinesNoMarkers" \| "AreaStacked" \| "AreaStacked100" \| "3DAreaStacked" \| "3DAreaStacked100" \| "DoughnutExploded" \| "RadarMarkers" \| "RadarFilled" \| "Surface" \| "SurfaceWireframe" \| "SurfaceTopView" \| "SurfaceTopViewWireframe" \| "Bubble" \| "Bubble3DEffect" \| "StockHLC" \| "StockOHLC" \| "StockVHLC" \| "StockVOHLC" \| "CylinderColClustered" \| "CylinderColStacked" \| "CylinderColStacked100" \| "CylinderBarClustered" \| "CylinderBarStacked" \| "CylinderBarStacked100" \| "CylinderCol" \| "ConeColClustered" \| "ConeColStacked" \| "ConeColStacked100" \| "ConeBarClustered" \| "ConeBarStacked" \| "ConeBarStacked100" \| "ConeCol" \| "PyramidColClustered" \| "PyramidColStacked" \| "PyramidColStacked100" \| "PyramidBarClustered" \| "PyramidBarStacked" \| "PyramidBarStacked100" \| "PyramidCol" \| "3DColumn" \| "Line" \| "3DLine" \| "3DPie" \| "Pie" \| "XYScatter" \| "3DArea" \| "Area" \| "Doughnut" \| "Radar" \| "Histogram" \| "Boxwhisker" \| "Pareto" \| "RegionMap" \| "Treemap" \| "Waterfall" \| "Sunburst" \| "Funnel", sourceData: Range, seriesBy?: "Auto" \| "Columns" \| "Rows")](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Creates a new chart.|
||[add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Creates a new chart.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Gets a chart based on its position in the collection.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Returns the number of charts in the worksheet. Read-only.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Gets the loaded child items in this collection.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[$all](/javascript/api/excel/excel.chartcollectionloadoptions#$all)||
||[axes](/javascript/api/excel/excel.chartcollectionloadoptions#axes)|For EACH ITEM in the collection: Represents chart axes.|
||[dataLabels](/javascript/api/excel/excel.chartcollectionloadoptions#datalabels)|For EACH ITEM in the collection: Represents the datalabels on the chart.|
||[format](/javascript/api/excel/excel.chartcollectionloadoptions#format)|For EACH ITEM in the collection: Encapsulates the format properties for the chart area.|
||[height](/javascript/api/excel/excel.chartcollectionloadoptions#height)|For EACH ITEM in the collection: Represents the height, in points, of the chart object.|
||[left](/javascript/api/excel/excel.chartcollectionloadoptions#left)|For EACH ITEM in the collection: The distance, in points, from the left side of the chart to the worksheet origin.|
||[legend](/javascript/api/excel/excel.chartcollectionloadoptions#legend)|For EACH ITEM in the collection: Represents the legend for the chart.|
||[name](/javascript/api/excel/excel.chartcollectionloadoptions#name)|For EACH ITEM in the collection: Represents the name of a chart object.|
||[series](/javascript/api/excel/excel.chartcollectionloadoptions#series)|For EACH ITEM in the collection: Represents either a single series or collection of series in the chart.|
||[title](/javascript/api/excel/excel.chartcollectionloadoptions#title)|For EACH ITEM in the collection: Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.|
||[top](/javascript/api/excel/excel.chartcollectionloadoptions#top)|For EACH ITEM in the collection: Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/javascript/api/excel/excel.chartcollectionloadoptions#width)|For EACH ITEM in the collection: Represents the width, in points, of the chart object.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[axes](/javascript/api/excel/excel.chartdata#axes)|Represents chart axes. Read-only.|
||[dataLabels](/javascript/api/excel/excel.chartdata#datalabels)|Represents the datalabels on the chart. Read-only.|
||[format](/javascript/api/excel/excel.chartdata#format)|Encapsulates the format properties for the chart area. Read-only.|
||[height](/javascript/api/excel/excel.chartdata#height)|Represents the height, in points, of the chart object.|
||[left](/javascript/api/excel/excel.chartdata#left)|The distance, in points, from the left side of the chart to the worksheet origin.|
||[legend](/javascript/api/excel/excel.chartdata#legend)|Represents the legend for the chart. Read-only.|
||[name](/javascript/api/excel/excel.chartdata#name)|Represents the name of a chart object.|
||[series](/javascript/api/excel/excel.chartdata#series)|Represents either a single series or collection of series in the chart. Read-only.|
||[title](/javascript/api/excel/excel.chartdata#title)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title. Read-only.|
||[top](/javascript/api/excel/excel.chartdata#top)|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/javascript/api/excel/excel.chartdata#width)|Represents the width, in points, of the chart object.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Represents the fill format of the current chart data label. Read-only.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.|
||[set(properties: Excel.ChartDataLabelFormat)](/javascript/api/excel/excel.chartdatalabelformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartDataLabelFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartdatalabelformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[font](/javascript/api/excel/excel.chartdatalabelformatdata#font)|Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartdatalabelformatloadoptions#font)|Represents the font attributes (font name, font size, color, etc.) for a chart data label.|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[font](/javascript/api/excel/excel.chartdatalabelformatupdatedata#font)|Represents the font attributes (font name, font size, color, etc.) for a chart data label.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Represents the format of chart data labels, which includes fill and font formatting. Read-only.|
||[separator](/javascript/api/excel/excel.chartdatalabels#separator)|String representing the separator used for the data labels on a chart.|
||[set(properties: Excel.ChartDataLabels)](/javascript/api/excel/excel.chartdatalabels#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartDataLabelsUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartdatalabels#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[format](/javascript/api/excel/excel.chartdatalabelsdata#format)|Represents the format of chart data labels, which includes fill and font formatting. Read-only.|
||[position](/javascript/api/excel/excel.chartdatalabelsdata#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[separator](/javascript/api/excel/excel.chartdatalabelsdata#separator)|String representing the separator used for the data labels on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsdata#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsdata#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsdata#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsdata#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsdata#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsdata#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelsloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartdatalabelsloadoptions#format)|Represents the format of chart data labels, which includes fill and font formatting.|
||[position](/javascript/api/excel/excel.chartdatalabelsloadoptions#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[separator](/javascript/api/excel/excel.chartdatalabelsloadoptions#separator)|String representing the separator used for the data labels on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsloadoptions#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsloadoptions#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsloadoptions#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsloadoptions#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[format](/javascript/api/excel/excel.chartdatalabelsupdatedata#format)|Represents the format of chart data labels, which includes fill and font formatting.|
||[position](/javascript/api/excel/excel.chartdatalabelsupdatedata#position)|DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.|
||[separator](/javascript/api/excel/excel.chartdatalabelsupdatedata#separator)|String representing the separator used for the data labels on a chart.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsupdatedata#showbubblesize)|Boolean value representing if the data label bubble size is visible or not.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showcategoryname)|Boolean value representing if the data label category name is visible or not.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsupdatedata#showlegendkey)|Boolean value representing if the data label legend key is visible or not.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsupdatedata#showpercentage)|Boolean value representing if the data label percentage is visible or not.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showseriesname)|Boolean value representing if the data label series name is visible or not.|
||[showValue](/javascript/api/excel/excel.chartdatalabelsupdatedata#showvalue)|Boolean value representing if the data label value is visible or not.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Clear the fill color of a chart element.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Sets the fill formatting of a chart element to a uniform color.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.chartfont#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.chartfont#name)|Font name (e.g. "Calibri")|
||[set(properties: Excel.ChartFont)](/javascript/api/excel/excel.chartfont#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartFontUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartfont#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[size](/javascript/api/excel/excel.chartfont#size)|Size of the font (e.g. 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.|
|[ChartFontData](/javascript/api/excel/excel.chartfontdata)|[bold](/javascript/api/excel/excel.chartfontdata#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.chartfontdata#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.chartfontdata#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.chartfontdata#name)|Font name (e.g. "Calibri")|
||[size](/javascript/api/excel/excel.chartfontdata#size)|Size of the font (e.g. 11)|
||[underline](/javascript/api/excel/excel.chartfontdata#underline)|Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.|
|[ChartFontLoadOptions](/javascript/api/excel/excel.chartfontloadoptions)|[$all](/javascript/api/excel/excel.chartfontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.chartfontloadoptions#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.chartfontloadoptions#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.chartfontloadoptions#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.chartfontloadoptions#name)|Font name (e.g. "Calibri")|
||[size](/javascript/api/excel/excel.chartfontloadoptions#size)|Size of the font (e.g. 11)|
||[underline](/javascript/api/excel/excel.chartfontloadoptions#underline)|Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.|
|[ChartFontUpdateData](/javascript/api/excel/excel.chartfontupdatedata)|[bold](/javascript/api/excel/excel.chartfontupdatedata#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.chartfontupdatedata#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.chartfontupdatedata#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.chartfontupdatedata#name)|Font name (e.g. "Calibri")|
||[size](/javascript/api/excel/excel.chartfontupdatedata#size)|Size of the font (e.g. 11)|
||[underline](/javascript/api/excel/excel.chartfontupdatedata#underline)|Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Represents the formatting of chart gridlines. Read-only.|
||[set(properties: Excel.ChartGridlines)](/javascript/api/excel/excel.chartgridlines#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartGridlinesUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartgridlines#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Boolean value representing if the axis gridlines are visible or not.|
|[ChartGridlinesData](/javascript/api/excel/excel.chartgridlinesdata)|[format](/javascript/api/excel/excel.chartgridlinesdata#format)|Represents the formatting of chart gridlines. Read-only.|
||[visible](/javascript/api/excel/excel.chartgridlinesdata#visible)|Boolean value representing if the axis gridlines are visible or not.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Represents chart line formatting. Read-only.|
||[set(properties: Excel.ChartGridlinesFormat)](/javascript/api/excel/excel.chartgridlinesformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartGridlinesFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartgridlinesformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartGridlinesFormatData](/javascript/api/excel/excel.chartgridlinesformatdata)|[line](/javascript/api/excel/excel.chartgridlinesformatdata#line)|Represents chart line formatting. Read-only.|
|[ChartGridlinesFormatLoadOptions](/javascript/api/excel/excel.chartgridlinesformatloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartgridlinesformatloadoptions#line)|Represents chart line formatting.|
|[ChartGridlinesFormatUpdateData](/javascript/api/excel/excel.chartgridlinesformatupdatedata)|[line](/javascript/api/excel/excel.chartgridlinesformatupdatedata#line)|Represents chart line formatting.|
|[ChartGridlinesLoadOptions](/javascript/api/excel/excel.chartgridlinesloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartgridlinesloadoptions#format)|Represents the formatting of chart gridlines.|
||[visible](/javascript/api/excel/excel.chartgridlinesloadoptions#visible)|Boolean value representing if the axis gridlines are visible or not.|
|[ChartGridlinesUpdateData](/javascript/api/excel/excel.chartgridlinesupdatedata)|[format](/javascript/api/excel/excel.chartgridlinesupdatedata#format)|Represents the formatting of chart gridlines.|
||[visible](/javascript/api/excel/excel.chartgridlinesupdatedata#visible)|Boolean value representing if the axis gridlines are visible or not.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Boolean value for whether the chart legend should overlap with the main body of the chart.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.|
||[set(properties: Excel.ChartLegend)](/javascript/api/excel/excel.chartlegend#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartLegendUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartlegend#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|A boolean value the represents the visibility of a ChartLegend object.|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[format](/javascript/api/excel/excel.chartlegenddata#format)|Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.|
||[overlay](/javascript/api/excel/excel.chartlegenddata#overlay)|Boolean value for whether the chart legend should overlap with the main body of the chart.|
||[position](/javascript/api/excel/excel.chartlegenddata#position)|Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.|
||[visible](/javascript/api/excel/excel.chartlegenddata#visible)|A boolean value the represents the visibility of a ChartLegend object.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Represents the fill format of an object, which includes background formatting information. Read-only.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.|
||[set(properties: Excel.ChartLegendFormat)](/javascript/api/excel/excel.chartlegendformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartLegendFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartlegendformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[font](/javascript/api/excel/excel.chartlegendformatdata#font)|Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[$all](/javascript/api/excel/excel.chartlegendformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartlegendformatloadoptions#font)|Represents the font attributes such as font name, font size, color, etc. of a chart legend.|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[font](/javascript/api/excel/excel.chartlegendformatupdatedata#font)|Represents the font attributes such as font name, font size, color, etc. of a chart legend.|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[$all](/javascript/api/excel/excel.chartlegendloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartlegendloadoptions#format)|Represents the formatting of a chart legend, which includes fill and font formatting.|
||[overlay](/javascript/api/excel/excel.chartlegendloadoptions#overlay)|Boolean value for whether the chart legend should overlap with the main body of the chart.|
||[position](/javascript/api/excel/excel.chartlegendloadoptions#position)|Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.|
||[visible](/javascript/api/excel/excel.chartlegendloadoptions#visible)|A boolean value the represents the visibility of a ChartLegend object.|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[format](/javascript/api/excel/excel.chartlegendupdatedata#format)|Represents the formatting of a chart legend, which includes fill and font formatting.|
||[overlay](/javascript/api/excel/excel.chartlegendupdatedata#overlay)|Boolean value for whether the chart legend should overlap with the main body of the chart.|
||[position](/javascript/api/excel/excel.chartlegendupdatedata#position)|Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.|
||[visible](/javascript/api/excel/excel.chartlegendupdatedata#visible)|A boolean value the represents the visibility of a ChartLegend object.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Clear the line format of a chart element.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|HTML color code representing the color of lines in the chart.|
||[set(properties: Excel.ChartLineFormat)](/javascript/api/excel/excel.chartlineformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartLineFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartlineformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[color](/javascript/api/excel/excel.chartlineformatdata#color)|HTML color code representing the color of lines in the chart.|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[$all](/javascript/api/excel/excel.chartlineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartlineformatloadoptions#color)|HTML color code representing the color of lines in the chart.|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[color](/javascript/api/excel/excel.chartlineformatupdatedata#color)|HTML color code representing the color of lines in the chart.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[$all](/javascript/api/excel/excel.chartloadoptions#$all)||
||[axes](/javascript/api/excel/excel.chartloadoptions#axes)|Represents chart axes.|
||[dataLabels](/javascript/api/excel/excel.chartloadoptions#datalabels)|Represents the datalabels on the chart.|
||[format](/javascript/api/excel/excel.chartloadoptions#format)|Encapsulates the format properties for the chart area.|
||[height](/javascript/api/excel/excel.chartloadoptions#height)|Represents the height, in points, of the chart object.|
||[left](/javascript/api/excel/excel.chartloadoptions#left)|The distance, in points, from the left side of the chart to the worksheet origin.|
||[legend](/javascript/api/excel/excel.chartloadoptions#legend)|Represents the legend for the chart.|
||[name](/javascript/api/excel/excel.chartloadoptions#name)|Represents the name of a chart object.|
||[series](/javascript/api/excel/excel.chartloadoptions#series)|Represents either a single series or collection of series in the chart.|
||[title](/javascript/api/excel/excel.chartloadoptions#title)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.|
||[top](/javascript/api/excel/excel.chartloadoptions#top)|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/javascript/api/excel/excel.chartloadoptions#width)|Represents the width, in points, of the chart object.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Encapsulates the format properties chart point. Read-only.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Returns the value of a chart point. Read-only.|
||[set(properties: Excel.ChartPoint)](/javascript/api/excel/excel.chartpoint#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartPointUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartpoint#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[format](/javascript/api/excel/excel.chartpointdata#format)|Encapsulates the format properties chart point. Read-only.|
||[value](/javascript/api/excel/excel.chartpointdata#value)|Returns the value of a chart point. Read-only.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Represents the fill format of a chart, which includes background formatting information. Read-only.|
||[set(properties: Excel.ChartPointFormat)](/javascript/api/excel/excel.chartpointformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartPointFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartpointformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[$all](/javascript/api/excel/excel.chartpointformatloadoptions#$all)||
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[$all](/javascript/api/excel/excel.chartpointloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointloadoptions#format)|Encapsulates the format properties chart point.|
||[value](/javascript/api/excel/excel.chartpointloadoptions#value)|Returns the value of a chart point. Read-only.|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[format](/javascript/api/excel/excel.chartpointupdatedata#format)|Encapsulates the format properties chart point.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Retrieve a point based on its position within the series.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Returns the number of chart points in the series. Read-only.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Gets the loaded child items in this collection.|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[$all](/javascript/api/excel/excel.chartpointscollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointscollectionloadoptions#format)|For EACH ITEM in the collection: Encapsulates the format properties chart point.|
||[value](/javascript/api/excel/excel.chartpointscollectionloadoptions#value)|For EACH ITEM in the collection: Returns the value of a chart point. Read-only.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Represents the name of a series in a chart.|
||[format](/javascript/api/excel/excel.chartseries#format)|Represents the formatting of a chart series, which includes fill and line formatting. Read-only.|
||[points](/javascript/api/excel/excel.chartseries#points)|Represents a collection of all points in the series. Read-only.|
||[set(properties: Excel.ChartSeries)](/javascript/api/excel/excel.chartseries#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartSeriesUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartseries#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Retrieves a series based on its position in the collection.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Returns the number of series in the collection. Read-only.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Gets the loaded child items in this collection.|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[$all](/javascript/api/excel/excel.chartseriescollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriescollectionloadoptions#format)|For EACH ITEM in the collection: Represents the formatting of a chart series, which includes fill and line formatting.|
||[name](/javascript/api/excel/excel.chartseriescollectionloadoptions#name)|For EACH ITEM in the collection: Represents the name of a series in a chart.|
||[points](/javascript/api/excel/excel.chartseriescollectionloadoptions#points)|For EACH ITEM in the collection: Represents a collection of all points in the series.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[format](/javascript/api/excel/excel.chartseriesdata#format)|Represents the formatting of a chart series, which includes fill and line formatting. Read-only.|
||[name](/javascript/api/excel/excel.chartseriesdata#name)|Represents the name of a series in a chart.|
||[points](/javascript/api/excel/excel.chartseriesdata#points)|Represents a collection of all points in the series. Read-only.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Represents the fill format of a chart series, which includes background formatting information. Read-only.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Represents line formatting. Read-only.|
||[set(properties: Excel.ChartSeriesFormat)](/javascript/api/excel/excel.chartseriesformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartSeriesFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartseriesformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartSeriesFormatData](/javascript/api/excel/excel.chartseriesformatdata)|[line](/javascript/api/excel/excel.chartseriesformatdata#line)|Represents line formatting. Read-only.|
|[ChartSeriesFormatLoadOptions](/javascript/api/excel/excel.chartseriesformatloadoptions)|[$all](/javascript/api/excel/excel.chartseriesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartseriesformatloadoptions#line)|Represents line formatting.|
|[ChartSeriesFormatUpdateData](/javascript/api/excel/excel.chartseriesformatupdatedata)|[line](/javascript/api/excel/excel.chartseriesformatupdatedata#line)|Represents line formatting.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[$all](/javascript/api/excel/excel.chartseriesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriesloadoptions#format)|Represents the formatting of a chart series, which includes fill and line formatting.|
||[name](/javascript/api/excel/excel.chartseriesloadoptions#name)|Represents the name of a series in a chart.|
||[points](/javascript/api/excel/excel.chartseriesloadoptions#points)|Represents a collection of all points in the series.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[format](/javascript/api/excel/excel.chartseriesupdatedata#format)|Represents the formatting of a chart series, which includes fill and line formatting.|
||[name](/javascript/api/excel/excel.chartseriesupdatedata#name)|Represents the name of a series in a chart.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|Boolean value representing if the chart title will overlay the chart or not.|
||[format](/javascript/api/excel/excel.charttitle#format)|Represents the formatting of a chart title, which includes fill and font formatting. Read-only.|
||[set(properties: Excel.ChartTitle)](/javascript/api/excel/excel.charttitle#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartTitleUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.charttitle#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[text](/javascript/api/excel/excel.charttitle#text)|Represents the title text of a chart.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|A boolean value the represents the visibility of a chart title object.|
|[ChartTitleData](/javascript/api/excel/excel.charttitledata)|[format](/javascript/api/excel/excel.charttitledata#format)|Represents the formatting of a chart title, which includes fill and font formatting. Read-only.|
||[overlay](/javascript/api/excel/excel.charttitledata#overlay)|Boolean value representing if the chart title will overlay the chart or not.|
||[text](/javascript/api/excel/excel.charttitledata#text)|Represents the title text of a chart.|
||[visible](/javascript/api/excel/excel.charttitledata#visible)|A boolean value the represents the visibility of a chart title object.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Represents the fill format of an object, which includes background formatting information. Read-only.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.|
||[set(properties: Excel.ChartTitleFormat)](/javascript/api/excel/excel.charttitleformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartTitleFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.charttitleformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartTitleFormatData](/javascript/api/excel/excel.charttitleformatdata)|[font](/javascript/api/excel/excel.charttitleformatdata#font)|Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.|
|[ChartTitleFormatLoadOptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[$all](/javascript/api/excel/excel.charttitleformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.charttitleformatloadoptions#font)|Represents the font attributes (font name, font size, color, etc.) for an object.|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[font](/javascript/api/excel/excel.charttitleformatupdatedata#font)|Represents the font attributes (font name, font size, color, etc.) for an object.|
|[ChartTitleLoadOptions](/javascript/api/excel/excel.charttitleloadoptions)|[$all](/javascript/api/excel/excel.charttitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttitleloadoptions#format)|Represents the formatting of a chart title, which includes fill and font formatting.|
||[overlay](/javascript/api/excel/excel.charttitleloadoptions#overlay)|Boolean value representing if the chart title will overlay the chart or not.|
||[text](/javascript/api/excel/excel.charttitleloadoptions#text)|Represents the title text of a chart.|
||[visible](/javascript/api/excel/excel.charttitleloadoptions#visible)|A boolean value the represents the visibility of a chart title object.|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[format](/javascript/api/excel/excel.charttitleupdatedata#format)|Represents the formatting of a chart title, which includes fill and font formatting.|
||[overlay](/javascript/api/excel/excel.charttitleupdatedata#overlay)|Boolean value representing if the chart title will overlay the chart or not.|
||[text](/javascript/api/excel/excel.charttitleupdatedata#text)|Represents the title text of a chart.|
||[visible](/javascript/api/excel/excel.charttitleupdatedata#visible)|A boolean value the represents the visibility of a chart title object.|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[axes](/javascript/api/excel/excel.chartupdatedata#axes)|Represents chart axes.|
||[dataLabels](/javascript/api/excel/excel.chartupdatedata#datalabels)|Represents the datalabels on the chart.|
||[format](/javascript/api/excel/excel.chartupdatedata#format)|Encapsulates the format properties for the chart area.|
||[height](/javascript/api/excel/excel.chartupdatedata#height)|Represents the height, in points, of the chart object.|
||[left](/javascript/api/excel/excel.chartupdatedata#left)|The distance, in points, from the left side of the chart to the worksheet origin.|
||[legend](/javascript/api/excel/excel.chartupdatedata#legend)|Represents the legend for the chart.|
||[name](/javascript/api/excel/excel.chartupdatedata#name)|Represents the name of a chart object.|
||[title](/javascript/api/excel/excel.chartupdatedata#title)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.|
||[top](/javascript/api/excel/excel.chartupdatedata#top)|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
||[width](/javascript/api/excel/excel.chartupdatedata#width)|Represents the width, in points, of the chart object.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.|
||[name](/javascript/api/excel/excel.nameditem#name)|The name of the object. Read-only.|
||[type](/javascript/api/excel/excel.nameditem#type)|Indicates the type of the value returned by the name's formula. See Excel.NamedItemType for details. Read-only.|
||[value](/javascript/api/excel/excel.nameditem#value)|Represents the value computed by the name's formula. For a named range, will return the range address. Read-only.|
||[set(properties: Excel.NamedItem)](/javascript/api/excel/excel.nameditem#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.NamedItemUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.nameditem#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Specifies whether the object is visible or not.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Gets a NamedItem object using its name.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Gets the loaded child items in this collection.|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[$all](/javascript/api/excel/excel.nameditemcollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemcollectionloadoptions#name)|For EACH ITEM in the collection: The name of the object. Read-only.|
||[type](/javascript/api/excel/excel.nameditemcollectionloadoptions#type)|For EACH ITEM in the collection: Indicates the type of the value returned by the name's formula. See Excel.NamedItemType for details. Read-only.|
||[value](/javascript/api/excel/excel.nameditemcollectionloadoptions#value)|For EACH ITEM in the collection: Represents the value computed by the name's formula. For a named range, will return the range address. Read-only.|
||[visible](/javascript/api/excel/excel.nameditemcollectionloadoptions#visible)|For EACH ITEM in the collection: Specifies whether the object is visible or not.|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[name](/javascript/api/excel/excel.nameditemdata#name)|The name of the object. Read-only.|
||[type](/javascript/api/excel/excel.nameditemdata#type)|Indicates the type of the value returned by the name's formula. See Excel.NamedItemType for details. Read-only.|
||[value](/javascript/api/excel/excel.nameditemdata#value)|Represents the value computed by the name's formula. For a named range, will return the range address. Read-only.|
||[visible](/javascript/api/excel/excel.nameditemdata#visible)|Specifies whether the object is visible or not.|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[$all](/javascript/api/excel/excel.nameditemloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemloadoptions#name)|The name of the object. Read-only.|
||[type](/javascript/api/excel/excel.nameditemloadoptions#type)|Indicates the type of the value returned by the name's formula. See Excel.NamedItemType for details. Read-only.|
||[value](/javascript/api/excel/excel.nameditemloadoptions#value)|Represents the value computed by the name's formula. For a named range, will return the range address. Read-only.|
||[visible](/javascript/api/excel/excel.nameditemloadoptions#visible)|Specifies whether the object is visible or not.|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[visible](/javascript/api/excel/excel.nameditemupdatedata#visible)|Specifies whether the object is visible or not.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.range#clear-applyto-)|Clear range values, format, fill, border, etc.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Clear range values, format, fill, border, etc.|
||[delete(shift: "Up" \| "Left")](/javascript/api/excel/excel.range#delete-shift-)|Deletes the cells associated with the range.|
||[delete(shift: Excel.DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|Deletes the cells associated with the range.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E15".|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Gets a column contained in the range.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Gets the range object that represents the rectangular intersection of the given ranges.|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".|
||[getLastColumn()](/javascript/api/excel/excel.range#getlastcolumn--)|Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".|
||[getLastRow()](/javascript/api/excel/excel.range#getlastrow--)|Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Gets a row contained in the range.|
||[insert(shift: "Down" \| "Right")](/javascript/api/excel/excel.range#insert-shift-)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.|
||[insert(shift: Excel.InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Represents Excel's number format code for the given range.|
||[address](/javascript/api/excel/excel.range#address)|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. "Sheet1!A1:B4"). Read-only.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Represents range reference for the specified range in the language of the user. Read-only.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|
||[columnCount](/javascript/api/excel/excel.range#columncount)|Represents the total number of columns in the range. Read-only.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|
||[format](/javascript/api/excel/excel.range#format)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.|
||[rowCount](/javascript/api/excel/excel.range#rowcount)|Returns the total number of rows in the range. Read-only.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|
||[text](/javascript/api/excel/excel.range#text)|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Represents the type of data of each cell. Read-only.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|The worksheet containing the current range. Read-only.|
||[select()](/javascript/api/excel/excel.range#select--)|Selects the specified range in the Excel UI.|
||[set(properties: Excel.Range)](/javascript/api/excel/excel.range#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RangeUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.range#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[track()](/javascript/api/excel/excel.range#track--)|Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.|
||[untrack()](/javascript/api/excel/excel.range#untrack--)|Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.|
||[values](/javascript/api/excel/excel.range#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.|
||[set(properties: Excel.RangeBorder)](/javascript/api/excel/excel.rangeborder#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RangeBorderUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.rangeborder#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[style](/javascript/api/excel/excel.rangeborder#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Specifies the weight of the border around a range. See Excel.BorderWeight for details.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem(index: "EdgeTop" \| "EdgeBottom" \| "EdgeLeft" \| "EdgeRight" \| "InsideVertical" \| "InsideHorizontal" \| "DiagonalDown" \| "DiagonalUp")](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Gets a border object using its name.|
||[getItem(index: Excel.BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Gets a border object using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Gets a border object using its index.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Number of border objects in the collection. Read-only.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Gets the loaded child items in this collection.|
|[RangeBorderCollectionLoadOptions](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.rangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangebordercollectionloadoptions#color)|For EACH ITEM in the collection: HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.rangebordercollectionloadoptions#sideindex)|For EACH ITEM in the collection: Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.|
||[style](/javascript/api/excel/excel.rangebordercollectionloadoptions#style)|For EACH ITEM in the collection: One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
||[weight](/javascript/api/excel/excel.rangebordercollectionloadoptions#weight)|For EACH ITEM in the collection: Specifies the weight of the border around a range. See Excel.BorderWeight for details.|
|[RangeBorderData](/javascript/api/excel/excel.rangeborderdata)|[color](/javascript/api/excel/excel.rangeborderdata#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.rangeborderdata#sideindex)|Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.|
||[style](/javascript/api/excel/excel.rangeborderdata#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
||[weight](/javascript/api/excel/excel.rangeborderdata#weight)|Specifies the weight of the border around a range. See Excel.BorderWeight for details.|
|[RangeBorderLoadOptions](/javascript/api/excel/excel.rangeborderloadoptions)|[$all](/javascript/api/excel/excel.rangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangeborderloadoptions#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.rangeborderloadoptions#sideindex)|Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.|
||[style](/javascript/api/excel/excel.rangeborderloadoptions#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
||[weight](/javascript/api/excel/excel.rangeborderloadoptions#weight)|Specifies the weight of the border around a range. See Excel.BorderWeight for details.|
|[RangeBorderUpdateData](/javascript/api/excel/excel.rangeborderupdatedata)|[color](/javascript/api/excel/excel.rangeborderupdatedata#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[style](/javascript/api/excel/excel.rangeborderupdatedata#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
||[weight](/javascript/api/excel/excel.rangeborderupdatedata#weight)|Specifies the weight of the border around a range. See Excel.BorderWeight for details.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[address](/javascript/api/excel/excel.rangedata#address)|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. "Sheet1!A1:B4"). Read-only.|
||[addressLocal](/javascript/api/excel/excel.rangedata#addresslocal)|Represents range reference for the specified range in the language of the user. Read-only.|
||[cellCount](/javascript/api/excel/excel.rangedata#cellcount)|Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|
||[columnCount](/javascript/api/excel/excel.rangedata#columncount)|Represents the total number of columns in the range. Read-only.|
||[columnIndex](/javascript/api/excel/excel.rangedata#columnindex)|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|
||[format](/javascript/api/excel/excel.rangedata#format)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.|
||[formulas](/javascript/api/excel/excel.rangedata#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangedata#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[numberFormat](/javascript/api/excel/excel.rangedata#numberformat)|Represents Excel's number format code for the given range.|
||[rowCount](/javascript/api/excel/excel.rangedata#rowcount)|Returns the total number of rows in the range. Read-only.|
||[rowIndex](/javascript/api/excel/excel.rangedata#rowindex)|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|
||[text](/javascript/api/excel/excel.rangedata#text)|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.rangedata#valuetypes)|Represents the type of data of each cell. Read-only.|
||[values](/javascript/api/excel/excel.rangedata#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Resets the range background.|
||[color](/javascript/api/excel/excel.rangefill#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
||[set(properties: Excel.RangeFill)](/javascript/api/excel/excel.rangefill#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RangeFillUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.rangefill#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[RangeFillData](/javascript/api/excel/excel.rangefilldata)|[color](/javascript/api/excel/excel.rangefilldata#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
|[RangeFillLoadOptions](/javascript/api/excel/excel.rangefillloadoptions)|[$all](/javascript/api/excel/excel.rangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangefillloadoptions#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
|[RangeFillUpdateData](/javascript/api/excel/excel.rangefillupdatedata)|[color](/javascript/api/excel/excel.rangefillupdatedata#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.rangefont#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.rangefont#name)|Font name (e.g. "Calibri")|
||[set(properties: Excel.RangeFont)](/javascript/api/excel/excel.rangefont#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RangeFontUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.rangefont#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[size](/javascript/api/excel/excel.rangefont#size)|Font size.|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.|
|[RangeFontData](/javascript/api/excel/excel.rangefontdata)|[bold](/javascript/api/excel/excel.rangefontdata#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.rangefontdata#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.rangefontdata#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.rangefontdata#name)|Font name (e.g. "Calibri")|
||[size](/javascript/api/excel/excel.rangefontdata#size)|Font size.|
||[underline](/javascript/api/excel/excel.rangefontdata#underline)|Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.|
|[RangeFontLoadOptions](/javascript/api/excel/excel.rangefontloadoptions)|[$all](/javascript/api/excel/excel.rangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.rangefontloadoptions#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.rangefontloadoptions#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.rangefontloadoptions#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.rangefontloadoptions#name)|Font name (e.g. "Calibri")|
||[size](/javascript/api/excel/excel.rangefontloadoptions#size)|Font size.|
||[underline](/javascript/api/excel/excel.rangefontloadoptions#underline)|Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.|
|[RangeFontUpdateData](/javascript/api/excel/excel.rangefontupdatedata)|[bold](/javascript/api/excel/excel.rangefontupdatedata#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.rangefontupdatedata#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.rangefontupdatedata#italic)|Represents the italic status of the font.|
||[name](/javascript/api/excel/excel.rangefontupdatedata#name)|Font name (e.g. "Calibri")|
||[size](/javascript/api/excel/excel.rangefontupdatedata#size)|Font size.|
||[underline](/javascript/api/excel/excel.rangefontupdatedata#underline)|Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|Collection of border objects that apply to the overall range. Read-only.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Returns the fill object defined on the overall range. Read-only.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Returns the font object defined on the overall range. Read-only.|
||[set(properties: Excel.RangeFormat)](/javascript/api/excel/excel.rangeformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RangeFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.rangeformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[borders](/javascript/api/excel/excel.rangeformatdata#borders)|Collection of border objects that apply to the overall range. Read-only.|
||[fill](/javascript/api/excel/excel.rangeformatdata#fill)|Returns the fill object defined on the overall range. Read-only.|
||[font](/javascript/api/excel/excel.rangeformatdata#font)|Returns the font object defined on the overall range. Read-only.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatdata#horizontalalignment)|Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatdata#verticalalignment)|Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.rangeformatdata#wraptext)|Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[$all](/javascript/api/excel/excel.rangeformatloadoptions#$all)||
||[borders](/javascript/api/excel/excel.rangeformatloadoptions#borders)|Collection of border objects that apply to the overall range.|
||[fill](/javascript/api/excel/excel.rangeformatloadoptions#fill)|Returns the fill object defined on the overall range.|
||[font](/javascript/api/excel/excel.rangeformatloadoptions#font)|Returns the font object defined on the overall range.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#horizontalalignment)|Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#verticalalignment)|Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.rangeformatloadoptions#wraptext)|Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[borders](/javascript/api/excel/excel.rangeformatupdatedata#borders)|Collection of border objects that apply to the overall range.|
||[fill](/javascript/api/excel/excel.rangeformatupdatedata#fill)|Returns the fill object defined on the overall range.|
||[font](/javascript/api/excel/excel.rangeformatupdatedata#font)|Returns the font object defined on the overall range.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#horizontalalignment)|Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#verticalalignment)|Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.|
||[wrapText](/javascript/api/excel/excel.rangeformatupdatedata#wraptext)|Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[$all](/javascript/api/excel/excel.rangeloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeloadoptions#address)|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. "Sheet1!A1:B4"). Read-only.|
||[addressLocal](/javascript/api/excel/excel.rangeloadoptions#addresslocal)|Represents range reference for the specified range in the language of the user. Read-only.|
||[cellCount](/javascript/api/excel/excel.rangeloadoptions#cellcount)|Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|
||[columnCount](/javascript/api/excel/excel.rangeloadoptions#columncount)|Represents the total number of columns in the range. Read-only.|
||[columnIndex](/javascript/api/excel/excel.rangeloadoptions#columnindex)|Represents the column number of the first cell in the range. Zero-indexed. Read-only.|
||[format](/javascript/api/excel/excel.rangeloadoptions#format)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.|
||[formulas](/javascript/api/excel/excel.rangeloadoptions#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeloadoptions#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[numberFormat](/javascript/api/excel/excel.rangeloadoptions#numberformat)|Represents Excel's number format code for the given range.|
||[rowCount](/javascript/api/excel/excel.rangeloadoptions#rowcount)|Returns the total number of rows in the range. Read-only.|
||[rowIndex](/javascript/api/excel/excel.rangeloadoptions#rowindex)|Returns the row number of the first cell in the range. Zero-indexed. Read-only.|
||[text](/javascript/api/excel/excel.rangeloadoptions#text)|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.rangeloadoptions#valuetypes)|Represents the type of data of each cell. Read-only.|
||[values](/javascript/api/excel/excel.rangeloadoptions#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
||[worksheet](/javascript/api/excel/excel.rangeloadoptions#worksheet)|The worksheet containing the current range.|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[format](/javascript/api/excel/excel.rangeupdatedata#format)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.|
||[formulas](/javascript/api/excel/excel.rangeupdatedata#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeupdatedata#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[numberFormat](/javascript/api/excel/excel.rangeupdatedata#numberformat)|Represents Excel's number format code for the given range.|
||[values](/javascript/api/excel/excel.rangeupdatedata#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Deletes the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|Gets the range object associated with the data body of the table.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|Gets the range object associated with header row of the table.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Gets the range object associated with the entire table.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|Gets the range object associated with totals row of the table.|
||[name](/javascript/api/excel/excel.table#name)|Name of the table.|
||[columns](/javascript/api/excel/excel.table#columns)|Represents a collection of all the columns in the table. Read-only.|
||[id](/javascript/api/excel/excel.table#id)|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
||[rows](/javascript/api/excel/excel.table#rows)|Represents a collection of all the rows in the table. Read-only.|
||[set(properties: Excel.Table)](/javascript/api/excel/excel.table#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TableUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.table#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
||[style](/javascript/api/excel/excel.table#style)|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Gets a table by Name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Gets a table based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Returns the number of tables in the workbook. Read-only.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Gets the loaded child items in this collection.|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecollectionloadoptions#$all)||
||[columns](/javascript/api/excel/excel.tablecollectionloadoptions#columns)|For EACH ITEM in the collection: Represents a collection of all the columns in the table.|
||[id](/javascript/api/excel/excel.tablecollectionloadoptions#id)|For EACH ITEM in the collection: Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
||[name](/javascript/api/excel/excel.tablecollectionloadoptions#name)|For EACH ITEM in the collection: Name of the table.|
||[rows](/javascript/api/excel/excel.tablecollectionloadoptions#rows)|For EACH ITEM in the collection: Represents a collection of all the rows in the table.|
||[showHeaders](/javascript/api/excel/excel.tablecollectionloadoptions#showheaders)|For EACH ITEM in the collection: Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
||[showTotals](/javascript/api/excel/excel.tablecollectionloadoptions#showtotals)|For EACH ITEM in the collection: Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
||[style](/javascript/api/excel/excel.tablecollectionloadoptions#style)|For EACH ITEM in the collection: Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Deletes the column from the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Gets the range object associated with the data body of the column.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Gets the range object associated with the header row of the column.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Gets the range object associated with the entire column.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Gets the range object associated with the totals row of the column.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Represents the name of the table column.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Returns a unique key that identifies the column within the table. Read-only.|
||[index](/javascript/api/excel/excel.tablecolumn#index)|Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|
||[set(properties: Excel.TableColumn)](/javascript/api/excel/excel.tablecolumn#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TableColumnUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.tablecolumn#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Adds a new column to the table.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Gets a column object by Name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Gets a column based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Returns the number of columns in the table. Read-only.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Gets the loaded child items in this collection.|
|[TableColumnCollectionLoadOptions](/javascript/api/excel/excel.tablecolumncollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecolumncollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumncollectionloadoptions#id)|For EACH ITEM in the collection: Returns a unique key that identifies the column within the table. Read-only.|
||[index](/javascript/api/excel/excel.tablecolumncollectionloadoptions#index)|For EACH ITEM in the collection: Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|
||[name](/javascript/api/excel/excel.tablecolumncollectionloadoptions#name)|For EACH ITEM in the collection: Represents the name of the table column.|
||[values](/javascript/api/excel/excel.tablecolumncollectionloadoptions#values)|For EACH ITEM in the collection: Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableColumnData](/javascript/api/excel/excel.tablecolumndata)|[id](/javascript/api/excel/excel.tablecolumndata#id)|Returns a unique key that identifies the column within the table. Read-only.|
||[index](/javascript/api/excel/excel.tablecolumndata#index)|Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|
||[name](/javascript/api/excel/excel.tablecolumndata#name)|Represents the name of the table column.|
||[values](/javascript/api/excel/excel.tablecolumndata#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableColumnLoadOptions](/javascript/api/excel/excel.tablecolumnloadoptions)|[$all](/javascript/api/excel/excel.tablecolumnloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumnloadoptions#id)|Returns a unique key that identifies the column within the table. Read-only.|
||[index](/javascript/api/excel/excel.tablecolumnloadoptions#index)|Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.|
||[name](/javascript/api/excel/excel.tablecolumnloadoptions#name)|Represents the name of the table column.|
||[values](/javascript/api/excel/excel.tablecolumnloadoptions#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableColumnUpdateData](/javascript/api/excel/excel.tablecolumnupdatedata)|[name](/javascript/api/excel/excel.tablecolumnupdatedata#name)|Represents the name of the table column.|
||[values](/javascript/api/excel/excel.tablecolumnupdatedata#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableData](/javascript/api/excel/excel.tabledata)|[columns](/javascript/api/excel/excel.tabledata#columns)|Represents a collection of all the columns in the table. Read-only.|
||[id](/javascript/api/excel/excel.tabledata#id)|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
||[name](/javascript/api/excel/excel.tabledata#name)|Name of the table.|
||[rows](/javascript/api/excel/excel.tabledata#rows)|Represents a collection of all the rows in the table. Read-only.|
||[showHeaders](/javascript/api/excel/excel.tabledata#showheaders)|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
||[showTotals](/javascript/api/excel/excel.tabledata#showtotals)|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
||[style](/javascript/api/excel/excel.tabledata#style)|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[$all](/javascript/api/excel/excel.tableloadoptions#$all)||
||[columns](/javascript/api/excel/excel.tableloadoptions#columns)|Represents a collection of all the columns in the table.|
||[id](/javascript/api/excel/excel.tableloadoptions#id)|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
||[name](/javascript/api/excel/excel.tableloadoptions#name)|Name of the table.|
||[rows](/javascript/api/excel/excel.tableloadoptions#rows)|Represents a collection of all the rows in the table.|
||[showHeaders](/javascript/api/excel/excel.tableloadoptions#showheaders)|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
||[showTotals](/javascript/api/excel/excel.tableloadoptions#showtotals)|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
||[style](/javascript/api/excel/excel.tableloadoptions#style)|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Deletes the row from the table.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Returns the range object associated with the entire row.|
||[index](/javascript/api/excel/excel.tablerow#index)|Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.|
||[set(properties: Excel.TableRow)](/javascript/api/excel/excel.tablerow#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TableRowUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.tablerow#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[values](/javascript/api/excel/excel.tablerow#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Adds one or more rows to the table. The return object will be the top of the newly added row(s).|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Gets a row based on its position in the collection.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Returns the number of rows in the table. Read-only.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Gets the loaded child items in this collection.|
|[TableRowCollectionLoadOptions](/javascript/api/excel/excel.tablerowcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablerowcollectionloadoptions#$all)||
||[index](/javascript/api/excel/excel.tablerowcollectionloadoptions#index)|For EACH ITEM in the collection: Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.|
||[values](/javascript/api/excel/excel.tablerowcollectionloadoptions#values)|For EACH ITEM in the collection: Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableRowData](/javascript/api/excel/excel.tablerowdata)|[index](/javascript/api/excel/excel.tablerowdata#index)|Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.|
||[values](/javascript/api/excel/excel.tablerowdata#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableRowLoadOptions](/javascript/api/excel/excel.tablerowloadoptions)|[$all](/javascript/api/excel/excel.tablerowloadoptions#$all)||
||[index](/javascript/api/excel/excel.tablerowloadoptions#index)|Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.|
||[values](/javascript/api/excel/excel.tablerowloadoptions#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableRowUpdateData](/javascript/api/excel/excel.tablerowupdatedata)|[values](/javascript/api/excel/excel.tablerowupdatedata#values)|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[name](/javascript/api/excel/excel.tableupdatedata#name)|Name of the table.|
||[showHeaders](/javascript/api/excel/excel.tableupdatedata#showheaders)|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
||[showTotals](/javascript/api/excel/excel.tableupdatedata#showtotals)|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
||[style](/javascript/api/excel/excel.tableupdatedata#style)|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange()](/javascript/api/excel/excel.workbook#getselectedrange--)|Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.|
||[application](/javascript/api/excel/excel.workbook#application)|Represents the Excel application instance that contains this workbook. Read-only.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Represents a collection of bindings that are part of the workbook. Read-only.|
||[names](/javascript/api/excel/excel.workbook#names)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.|
||[tables](/javascript/api/excel/excel.workbook#tables)|Represents a collection of tables associated with the workbook. Read-only.|
||[worksheets](/javascript/api/excel/excel.workbook#worksheets)|Represents a collection of worksheets associated with the workbook. Read-only.|
||[set(properties: Excel.Workbook)](/javascript/api/excel/excel.workbook#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.WorkbookUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.workbook#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[bindings](/javascript/api/excel/excel.workbookdata#bindings)|Represents a collection of bindings that are part of the workbook. Read-only.|
||[names](/javascript/api/excel/excel.workbookdata#names)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.|
||[tables](/javascript/api/excel/excel.workbookdata#tables)|Represents a collection of tables associated with the workbook. Read-only.|
||[worksheets](/javascript/api/excel/excel.workbookdata#worksheets)|Represents a collection of worksheets associated with the workbook. Read-only.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[$all](/javascript/api/excel/excel.workbookloadoptions#$all)||
||[application](/javascript/api/excel/excel.workbookloadoptions#application)|Represents the Excel application instance that contains this workbook.|
||[bindings](/javascript/api/excel/excel.workbookloadoptions#bindings)|Represents a collection of bindings that are part of the workbook.|
||[tables](/javascript/api/excel/excel.workbookloadoptions#tables)|Represents a collection of tables associated with the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Activate the worksheet in the Excel UI.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Deletes the worksheet from the workbook. Note that if the worksheet's visibility is set to "VeryHidden", the delete operation will fail with a GeneralException.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#getrange-address-)|Gets the range object, representing a single rectangular block of cells, specified by the address or name.|
||[name](/javascript/api/excel/excel.worksheet#name)|The display name of the worksheet.|
||[position](/javascript/api/excel/excel.worksheet#position)|The zero-based position of the worksheet within the workbook.|
||[charts](/javascript/api/excel/excel.worksheet#charts)|Returns collection of charts that are part of the worksheet. Read-only.|
||[id](/javascript/api/excel/excel.worksheet#id)|Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|
||[tables](/javascript/api/excel/excel.worksheet#tables)|Collection of tables that are part of the worksheet. Read-only.|
||[set(properties: Excel.Worksheet)](/javascript/api/excel/excel.worksheet#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.WorksheetUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.worksheet#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|The Visibility of the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Gets the currently active worksheet in the workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Gets a worksheet object using its Name or ID.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Gets the loaded child items in this collection.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[$all](/javascript/api/excel/excel.worksheetcollectionloadoptions#$all)||
||[charts](/javascript/api/excel/excel.worksheetcollectionloadoptions#charts)|For EACH ITEM in the collection: Returns collection of charts that are part of the worksheet.|
||[id](/javascript/api/excel/excel.worksheetcollectionloadoptions#id)|For EACH ITEM in the collection: Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|
||[name](/javascript/api/excel/excel.worksheetcollectionloadoptions#name)|For EACH ITEM in the collection: The display name of the worksheet.|
||[position](/javascript/api/excel/excel.worksheetcollectionloadoptions#position)|For EACH ITEM in the collection: The zero-based position of the worksheet within the workbook.|
||[tables](/javascript/api/excel/excel.worksheetcollectionloadoptions#tables)|For EACH ITEM in the collection: Collection of tables that are part of the worksheet.|
||[visibility](/javascript/api/excel/excel.worksheetcollectionloadoptions#visibility)|For EACH ITEM in the collection: The Visibility of the worksheet.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[charts](/javascript/api/excel/excel.worksheetdata#charts)|Returns collection of charts that are part of the worksheet. Read-only.|
||[id](/javascript/api/excel/excel.worksheetdata#id)|Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|
||[name](/javascript/api/excel/excel.worksheetdata#name)|The display name of the worksheet.|
||[position](/javascript/api/excel/excel.worksheetdata#position)|The zero-based position of the worksheet within the workbook.|
||[tables](/javascript/api/excel/excel.worksheetdata#tables)|Collection of tables that are part of the worksheet. Read-only.|
||[visibility](/javascript/api/excel/excel.worksheetdata#visibility)|The Visibility of the worksheet.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[$all](/javascript/api/excel/excel.worksheetloadoptions#$all)||
||[charts](/javascript/api/excel/excel.worksheetloadoptions#charts)|Returns collection of charts that are part of the worksheet.|
||[id](/javascript/api/excel/excel.worksheetloadoptions#id)|Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|
||[name](/javascript/api/excel/excel.worksheetloadoptions#name)|The display name of the worksheet.|
||[position](/javascript/api/excel/excel.worksheetloadoptions#position)|The zero-based position of the worksheet within the workbook.|
||[tables](/javascript/api/excel/excel.worksheetloadoptions#tables)|Collection of tables that are part of the worksheet.|
||[visibility](/javascript/api/excel/excel.worksheetloadoptions#visibility)|The Visibility of the worksheet.|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[name](/javascript/api/excel/excel.worksheetupdatedata#name)|The display name of the worksheet.|
||[position](/javascript/api/excel/excel.worksheetupdatedata#position)|The zero-based position of the worksheet within the workbook.|
||[visibility](/javascript/api/excel/excel.worksheetupdatedata#visibility)|The Visibility of the worksheet.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel&view=excel-js-1.1)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
