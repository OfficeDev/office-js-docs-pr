---
title: Excel JavaScript API requirement set 1.9
description: 'Details about the ExcelApi 1.9 requirement set.'
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.9

More than 500 new Excel APIs were introduced with the 1.9 requirement set. The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Shapes](../../excel/excel-add-ins-shapes.md) | Insert, position, and format images, geometric shapes and text boxes. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Auto Filter](../../excel/excel-add-ins-worksheets.md#filter-data) | Add filters to ranges. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](../../excel/excel-add-ins-multiple-ranges.md) | Support for discontinuous ranges. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Special Cells](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Get cells containing dates, comments, or formulas within a range. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Find](../../excel/excel-add-ins-ranges-string-match.md) | Find values or formulas within a range or worksheet. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copy and Paste](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Copy values, formats, and formulas from one range to another. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](../../excel/performance.md#suspend-calculation-temporarily) | Greater control over the Excel calculation engine. | [Application](/javascript/api/excel/excel.application) |
| New Charts | Explore our new supported chart types: maps, box and whisker, waterfall, sunburst, pareto. and funnel. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | New capabilities with range formats. | [Range](/javascript/api/excel/excel.rangeformat) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.9. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.9 or earlier, see [Excel APIs in requirement set 1.9 or earlier](/javascript/api/excel?view=excel-js-1.9&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationEngineVersion)|Returns the Excel calculation engine version used for the last full recalculation.|
||[calculationState](/javascript/api/excel/excel.application#calculationState)|Returns the calculation state of the application.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativeCalculation)|Returns the iterative calculation settings.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendScreenUpdatingUntilNextSync__)|Suspends screen updating until the next `context.sync()` is called.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply_range__columnIndex__criteria_)|Applies the AutoFilter to a range.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearCriteria__)|Clears the filter criteria of the AutoFilter.|
||[getRange()](/javascript/api/excel/excel.autofilter#getRange__)|Returns the `Range` object that represents the range to which the AutoFilter applies.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getRangeOrNullObject__)|Returns the `Range` object that represents the range to which the AutoFilter applies.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|An array that holds all the filter criteria in the autofiltered range.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Specifies if the AutoFilter is enabled.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isDataFiltered)|Specifies if the AutoFilter has filter criteria.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply__)|Applies the specified Autofilter object currently on the range.|
||[remove()](/javascript/api/excel/excel.autofilter#remove__)|Removes the AutoFilter for the range.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Represents the `color` property of a single border.|
||[style](/javascript/api/excel/excel.cellborder#style)|Represents the `style` property of a single border.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintAndShade)|Represents the `tintAndShade` property of a single border.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Represents the `weight` property of a single border.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|Represents the `format.borders.bottom` property.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonalDown)|Represents the `format.borders.diagonalDown` property.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalUp)|Represents the `format.borders.diagonalUp` property.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Represents the `format.borders.horizontal` property.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Represents the `format.borders.left` property.|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|Represents the `format.borders.right` property.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Represents the `format.borders.top` property.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Represents the `format.borders.vertical` property.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addressLocal)|Represents the `addressLocal` property.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Represents the `hidden` property.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Represents the `format.fill.color` property.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Represents the `format.fill.pattern` property.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patternColor)|Represents the `format.fill.patternColor` property.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patternTintAndShade)|Represents the `format.fill.patternTintAndShade` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintAndShade)|Represents the `format.fill.tintAndShade` property.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Represents the `format.font.bold` property.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Represents the `format.font.color` property.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Represents the `format.font.italic` property.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Represents the `format.font.name` property.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Represents the `format.font.size` property.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Represents the `format.font.strikethrough` property.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Represents the `format.font.subscript` property.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Represents the `format.font.superscript` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintAndShade)|Represents the `format.font.tintAndShade` property.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Represents the `format.font.underline` property.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoIndent)|Represents the `autoIndent` property.|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Represents the `borders` property.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Represents the `fill` property.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|Represents the `font` property.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalAlignment)|Represents the `horizontalAlignment` property.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentLevel)|Represents the `indentLevel` property.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Represents the `protection` property.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingOrder)|Represents the `readingOrder` property.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinkToFit)|Represents the `shrinkToFit` property.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textOrientation)|Represents the `textOrientation` property.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)|Represents the `useStandardHeight` property.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth)|Represents the `useStandardWidth` property.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalAlignment)|Represents the `verticalAlignment` property.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wrapText)|Represents the `wrapText` property.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulaHidden)|Represents the `format.protection.formulaHidden` property.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Represents the `format.protection.locked` property.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueAfter)|Represents the value after the change.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valueBefore)|Represents the value before the change.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valueTypeAfter)|Represents the type of value after the change.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valueTypeBefore)|Represents the type of value before the change.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate__)|Activates the chart in the Excel UI.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotOptions)|Encapsulates the options for a pivot chart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorScheme)|Specifies the color scheme of the chart.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedCorners)|Specifies if the chart area of the chart has rounded corners.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linkNumberFormat)|Specifies if the number format is linked to the cells.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowOverflow)|Specifies if bin overflow is enabled in a histogram chart or pareto chart.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowUnderflow)|Specifies if bin underflow is enabled in a histogram chart or pareto chart.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Specifies the bin count of a histogram chart or pareto chart.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowValue)|Specifies the bin overflow value of a histogram chart or pareto chart.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Specifies the bin's type for a histogram chart or pareto chart.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowValue)|Specifies the bin underflow value of a histogram chart or pareto chart.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Specifies the bin width value of a histogram chart or pareto chart.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartileCalculation)|Specifies if the quartile calculation type of a box and whisker chart.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showInnerPoints)|Specifies if inner points are shown in a box and whisker chart.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanLine)|Specifies if the mean line is shown in a box and whisker chart.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanMarker)|Specifies if the mean marker is shown in a box and whisker chart.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showOutlierPoints)|Specifies if outlier points are shown in a box and whisker chart.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linkNumberFormat)|Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linkNumberFormat)|Specifies if the number format is linked to the cells.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endStyleCap)|Specifies if error bars have an end style cap.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Specifies which parts of the error bars to include.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Specifies the formatting type of the error bars.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|The type of range marked by the error bars.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Specifies whether the error bars are displayed.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Represents the chart line formatting.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelStrategy)|Specifies the series map labels strategy of a region map chart.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Specifies the series mapping level of a region map chart.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectionType)|Specifies the series projection type of a region map chart.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showAxisFieldButtons)|Specifies whether to display the axis field buttons on a PivotChart.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showLegendFieldButtons)|Specifies whether to display the legend field buttons on a PivotChart.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showReportFilterFieldButtons)|Specifies whether to display the report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showValueFieldButtons)|Specifies whether to display the show value field buttons on a PivotChart.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubbleScale)|This can be an integer value from 0 (zero) to 300, representing the percentage of the default size.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientMaximumColor)|Specifies the color for maximum value of a region map chart series.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientMaximumType)|Specifies the type for maximum value of a region map chart series.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientMaximumValue)|Specifies the maximum value of a region map chart series.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientMidpointColor)|Specifies the color for the midpoint value of a region map chart series.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientMidpointType)|Specifies the type for the midpoint value of a region map chart series.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientMidpointValue)|Specifies the midpoint value of a region map chart series.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientMinimumColor)|Specifies the color for the minimum value of a region map chart series.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientMinimumType)|Specifies the type for the minimum value of a region map chart series.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientMinimumValue)|Specifies the minimum value of a region map chart series.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientStyle)|Specifies the series gradient style of a region map chart.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertColor)|Specifies the fill color for negative data points in a series.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentLabelStrategy)|Specifies the series parent label strategy area for a treemap chart.|
||[binOptions](/javascript/api/excel/excel.chartseries#binOptions)|Encapsulates the bin options for histogram charts and pareto charts.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskerOptions)|Encapsulates the options for the box and whisker charts.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapOptions)|Encapsulates the options for a region map chart.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xErrorBars)|Represents the error bar object of a chart series.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yErrorBars)|Represents the error bar object of a chart series.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showConnectorLines)|Specifies whether connector lines are shown in waterfall charts.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showLeaderLines)|Specifies whether leader lines are displayed for each data label in the series.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitValue)|Specifies the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linkNumberFormat)|Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addressLocal)|Represents the `addressLocal` property.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnIndex)|Represents the `columnIndex` property.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getRanges__)|Returns the `RangeAreas`, comprising one or more rectangular ranges, to which the conditonal format is applied.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getInvalidCells__)|Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getInvalidCellsOrNullObject__)|Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subField)|The property used by the filter to do a rich filter on rich values.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Returns the shape identifier.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Returns the `Shape` object for the geometric shape.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getCount__)|Returns the number of shapes in the shape group.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getItem_key_)|Gets a shape using its name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getItemAt_index_)|Gets a shape based on its position in the collection.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Gets the loaded child items in this collection.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerFooter)|The center footer of the worksheet.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerHeader)|The center header of the worksheet.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftFooter)|The left footer of the worksheet.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftHeader)|The left header of the worksheet.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightFooter)|The right footer of the worksheet.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightHeader)|The right header of the worksheet.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultForAllPages)|The general header/footer, used for all pages unless even/odd or first page is specified.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenPages)|The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstPage)|The first page header/footer, for all other pages general or even/odd is used.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddPages)|The header/footer to use for odd pages, even header/footer needs to be specified for even pages.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|The state by which headers/footers are set.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#useSheetMargins)|Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#useSheetScale)|Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Returns the format of the image.|
||[id](/javascript/api/excel/excel.image#id)|Specifies the shape identifier for the image object.|
||[shape](/javascript/api/excel/excel.image#shape)|Returns the `Shape` object associated with the image.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxChange)|Specifies the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxIteration)|Specifies the maximum number of iterations that Excel can use to resolve a circular reference.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginArrowheadLength)|Represents the length of the arrowhead at the beginning of the specified line.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginArrowheadStyle)|Represents the style of the arrowhead at the beginning of the specified line.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginArrowheadWidth)|Represents the width of the arrowhead at the beginning of the specified line.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectBeginShape_shape__connectionSite_)|Attaches the beginning of the specified connector to a specified shape.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectEndShape_shape__connectionSite_)|Attaches the end of the specified connector to a specified shape.|
||[connectorType](/javascript/api/excel/excel.line#connectorType)|Represents the connector type for the line.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectBeginShape__)|Detaches the beginning of the specified connector from a shape.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectEndShape__)|Detaches the end of the specified connector from a shape.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endArrowheadLength)|Represents the length of the arrowhead at the end of the specified line.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endArrowheadStyle)|Represents the style of the arrowhead at the end of the specified line.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endArrowheadWidth)|Represents the width of the arrowhead at the end of the specified line.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginConnectedShape)|Represents the shape to which the beginning of the specified line is attached.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginConnectedSite)|Represents the connection site to which the beginning of a connector is connected.|
||[endConnectedShape](/javascript/api/excel/excel.line#endConnectedShape)|Represents the shape to which the end of the specified line is attached.|
||[endConnectedSite](/javascript/api/excel/excel.line#endConnectedSite)|Represents the connection site to which the end of a connector is connected.|
||[id](/javascript/api/excel/excel.line#id)|Specifies the shape identifier.|
||[isBeginConnected](/javascript/api/excel/excel.line#isBeginConnected)|Specifies if the beginning of the specified line is connected to a shape.|
||[isEndConnected](/javascript/api/excel/excel.line#isEndConnected)|Specifies if the end of the specified line is connected to a shape.|
||[shape](/javascript/api/excel/excel.line#shape)|Returns the `Shape` object associated with the line.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete__)|Deletes a page break object.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getCellAfterBreak__)|Gets the first cell after the page break.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnIndex)|Specifies the column index for the page break.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowIndex)|Specifies the row index for the page break.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add_pageBreakRange_)|Adds a page break before the top-left cell of the range specified.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getCount__)|Gets the number of page breaks in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getItem_index_)|Gets a page break object via the index.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Gets the loaded child items in this collection.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removePageBreaks__)|Resets all manual page breaks in the collection.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackAndWhite)|The worksheet's black and white print option.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottomMargin)|The worksheet's bottom page margin to use for printing in points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerHorizontally)|The worksheet's center horizontally flag.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centerVertically)|The worksheet's center vertically flag.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftMode)|The worksheet's draft mode option.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstPageNumber)|The worksheet's first page number to print.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footerMargin)|The worksheet's footer margin, in points, for use when printing.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getPrintArea__)|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintAreaOrNullObject__)|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumns__)|Gets the range object representing the title columns.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumnsOrNullObject__)|Gets the range object representing the title columns.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getPrintTitleRows__)|Gets the range object representing the title rows.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleRowsOrNullObject__)|Gets the range object representing the title rows.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headerMargin)|The worksheet's header margin, in points, for use when printing.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftMargin)|The worksheet's left margin, in points, for use when printing.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|The worksheet's orientation of the page.|
||[paperSize](/javascript/api/excel/excel.pagelayout#paperSize)|The worksheet's paper size of the page.|
||[printComments](/javascript/api/excel/excel.pagelayout#printComments)|Specifies if the worksheet's comments should be displayed when printing.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printErrors)|The worksheet's print errors option.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printGridlines)|Specifies if the worksheet's gridlines will be printed.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printHeadings)|Specifies if the worksheet's headings will be printed.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printOrder)|The worksheet's page print order option.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersFooters)|Header and footer configuration for the worksheet.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightMargin)|The worksheet's right margin, in points, for use when printing.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setPrintArea_printArea_)|Sets the worksheet's print area.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setPrintMargins_unit__marginOptions_)|Sets the worksheet's page margins with units.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleColumns_printTitleColumns_)|Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleRows_printTitleRows_)|Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topMargin)|The worksheet's top margin, in points, for use when printing.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|The worksheet's print zoom options.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Specifies the page layout bottom margin in the unit specified to use for printing.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Specifies the page layout footer margin in the unit specified to use for printing.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Specifies the page layout header margin in the unit specified to use for printing.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Specifies the page layout left margin in the unit specified to use for printing.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Specifies the page layout right margin in the unit specified to use for printing.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Specifies the page layout top margin in the unit specified to use for printing.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalFitToPages)|Number of pages to fit horizontally.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|Print page scale value can be between 10 and 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalFitToPages)|Number of pages to fit vertically.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortByValues_sortBy__valuesHierarchy__pivotItemScope_)|Sorts the PivotField by specified values in a given scope.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoFormat)|Specifies if formatting will be automatically formatted when it’s refreshed or when fields are moved.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getDataHierarchy_cell_)|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getPivotItems_axis__cell_)|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveFormatting)|Specifies if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setAutoSortOnCell_cell__sortBy_)|Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enableDataValueEditing)|Specifies if the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#useCustomSortLists)|Specifies if the PivotTable uses custom lists when sorting.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autoFill_destinationRange__autoFillType_)|Fills range from the current range to the destination range using the specified AutoFill logic.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertDataTypeToText__)|Converts the range cells with data types into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#convertToLinkedDataType_serviceID__languageCulture_)|Converts the range cells into linked data types in the worksheet.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|Copies cell data or formatting from the source range or `RangeAreas` to the current range.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find_text__criteria_)|Finds the given string based on the criteria specified.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findOrNullObject_text__criteria_)|Finds the given string based on the criteria specified.|
||[flashFill()](/javascript/api/excel/excel.range#flashFill__)|Does a Flash Fill to the current range.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getCellProperties_cellPropertiesLoadOptions_)|Returns a 2D array, encapsulating the data for each cell's font, fill, borders, alignment, and other properties.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getColumnProperties_columnPropertiesLoadOptions_)|Returns a single-dimensional array, encapsulating the data for each column's font, fill, borders, alignment, and other properties.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getRowProperties_rowPropertiesLoadOptions_)|Returns a single-dimensional array, encapsulating the data for each row's font, fill, borders, alignment, and other properties.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCells_cellType__cellValueType_)|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCellsOrNullObject_cellType__cellValueType_)|Gets the `RangeAreas` object, comprising one or more ranges, that represents all the cells that match the specified type and value.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getTables_fullyContained_)|Gets a scoped collection of tables that overlap with the range.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkedDataTypeState)|Represents the data type state of each cell.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_)|Removes duplicate values from the range specified by the columns.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceAll_text__replacement__criteria_)|Finds and replaces the given string based on the criteria specified within the current range.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setCellProperties_cellPropertiesData_)|Updates the range based on a 2D array of cell properties, encapsulating things like font, fill, borders, and alignment.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setColumnProperties_columnPropertiesData_)|Updates the range based on a single-dimensional array of column properties, encapsulating things like font, fill, borders, and alignment.|
||[setDirty()](/javascript/api/excel/excel.range#setDirty__)|Set a range to be recalculated when the next recalculation occurs.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setRowProperties_rowPropertiesData_)|Updates the range based on a single-dimensional array of row properties, encapsulating things like font, fill, borders, and alignment.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate__)|Calculates all cells in the `RangeAreas`.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear_applyTo_)|Clears values, format, fill, border, and other properties on each of the areas that comprise this `RangeAreas` object.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertDataTypeToText__)|Converts all cells in the `RangeAreas` with data types into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#convertToLinkedDataType_serviceID__languageCulture_)|Converts all cells in the `RangeAreas` into linked data types.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|Copies cell data or formatting from the source range or `RangeAreas` to the current `RangeAreas`.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getEntireColumn__)|Returns a `RangeAreas` object that represents the entire columns of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11, H2", it returns a `RangeAreas` that represents columns "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getEntireRow__)|Returns a `RangeAreas` object that represents the entire rows of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11", it returns a `RangeAreas` that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersection_anotherRange_)|Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersectionOrNullObject_anotherRange_)|Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getOffsetRangeAreas_rowOffset__columnOffset_)|Returns a `RangeAreas` object that is shifted by the specific row and column offset.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCells_cellType__cellValueType_)|Returns a `RangeAreas` object that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCellsOrNullObject_cellType__cellValueType_)|Returns a `RangeAreas` object that represents all the cells that match the specified type and value.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#getTables_fullyContained_)|Returns a scoped collection of tables that overlap with any range in this `RangeAreas` object.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreas_valuesOnly_)|Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreasOrNullObject_valuesOnly_)|Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Returns the `RangeAreas` reference in A1-style.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addressLocal)|Returns the `RangeAreas` reference in the user locale.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areaCount)|Returns the number of rectangular ranges that comprise this `RangeAreas` object.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Returns a collection of rectangular ranges that comprise this `RangeAreas` object.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellCount)|Returns the number of cells in the `RangeAreas` object, summing up the cell counts of all of the individual rectangular ranges.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalFormats)|Returns a collection of conditional formats that intersect with any cells in this `RangeAreas` object.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#dataValidation)|Returns a data validation object for all ranges in the `RangeAreas`.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Returns a `RangeFormat` object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the `RangeAreas` object.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isEntireColumn)|Specifies if all the ranges on this `RangeAreas` object represent entire columns (e.g., "A:C, Q:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isEntireRow)|Specifies if all the ranges on this `RangeAreas` object represent entire rows (e.g., "1:3, 5:7").|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Returns the worksheet for the current `RangeAreas`.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setDirty__)|Sets the `RangeAreas` to be recalculated when the next recalculation occurs.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Represents the style for all ranges in this `RangeAreas` object.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintAndShade)|Specifies a double that lightens or darkens a color for the range border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintAndShade)|Specifies a double that lightens or darkens a color for range borders.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getCount__)|Returns the number of ranges in the `RangeCollection`.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getItemAt_index_)|Returns the range object based on its position in the `RangeCollection`.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Gets the loaded child items in this collection.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|The pattern of a range.|
||[patternColor](/javascript/api/excel/excel.rangefill#patternColor)|The HTML color code representing the color of the range pattern, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patternTintAndShade)|Specifies a double that lightens or darkens a pattern color for the range fill.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintAndShade)|Specifies a double that lightens or darkens a color for the range fill.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Specifies the strikethrough status of font.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Specifies the subscript status of font.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Specifies the superscript status of font.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintAndShade)|Specifies a double that lightens or darkens a color for the range font.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoIndent)|Specifies if text is automatically indented when text alignment is set to equal distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentLevel)|An integer from 0 to 250 that indicates the indent level.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingOrder)|The reading order for the range.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinkToFit)|Specifies if text automatically shrinks to fit in the available column width.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Number of duplicated rows removed by the operation.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueRemaining)|Number of remaining unique rows present in the resulting range.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completeMatch)|Specifies if the match needs to be complete or partial.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchCase)|Specifies if the match is case-sensitive.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addressLocal)|Represents the `addressLocal` property.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowIndex)|Represents the `rowIndex` property.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completeMatch)|Specifies if the match needs to be complete or partial.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchCase)|Specifies if the match is case-sensitive.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchDirection)|Specifies the search direction.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Represents the `format` property.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Represents the `hyperlink` property.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Represents the `style` property.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnHidden)|Represents the `columnHidden` property.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnWidth)||
||[format: Excel.CellPropertiesFormat & {
            columnWidth?](/javascript/api/excel/excel.settablecolumnproperties#format)|Represents the `format` property.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel.CellPropertiesFormat & {
            rowHeight?](/javascript/api/excel/excel.settablerowproperties#format)|Represents the `format` property.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowHeight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowHidden)|Represents the `rowHidden` property.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#altTextDescription)|Specifies the alternative description text for a `Shape` object.|
||[altTextTitle](/javascript/api/excel/excel.shape#altTextTitle)|Specifies the alternative title text for a `Shape` object.|
||[delete()](/javascript/api/excel/excel.shape#delete__)|Removes the shape from the worksheet.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricShapeType)|Specifies the geometric shape type of this geometric shape.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getAsImage_format_)|Converts the shape to an image and returns the image as a base64-encoded string.|
||[height](/javascript/api/excel/excel.shape#height)|Specifies the height, in points, of the shape.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementLeft_increment_)|Moves the shape horizontally by the specified number of points.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementRotation_increment_)|Rotates the shape clockwise around the z-axis by the specified number of degrees.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementTop_increment_)|Moves the shape vertically by the specified number of points.|
||[left](/javascript/api/excel/excel.shape#left)|The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockAspectRatio)|Specifies if the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.shape#name)|Specifies the name of the shape.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionSiteCount)|Returns the number of connection sites on this shape.|
||[fill](/javascript/api/excel/excel.shape#fill)|Returns the fill formatting of this shape.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricShape)|Returns the geometric shape associated with the shape.|
||[group](/javascript/api/excel/excel.shape#group)|Returns the shape group associated with the shape.|
||[id](/javascript/api/excel/excel.shape#id)|Specifies the shape identifier.|
||[image](/javascript/api/excel/excel.shape#image)|Returns the image associated with the shape.|
||[level](/javascript/api/excel/excel.shape#level)|Specifies the level of the specified shape.|
||[line](/javascript/api/excel/excel.shape#line)|Returns the line associated with the shape.|
||[lineFormat](/javascript/api/excel/excel.shape#lineFormat)|Returns the line formatting of this shape.|
||[onActivated](/javascript/api/excel/excel.shape#onActivated)|Occurs when the shape is activated.|
||[onDeactivated](/javascript/api/excel/excel.shape#onDeactivated)|Occurs when the shape is deactivated.|
||[parentGroup](/javascript/api/excel/excel.shape#parentGroup)|Specifies the parent group of this shape.|
||[textFrame](/javascript/api/excel/excel.shape#textFrame)|Returns the text frame object of this shape.|
||[type](/javascript/api/excel/excel.shape#type)|Returns the type of this shape.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zOrderPosition)|Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Specifies the rotation, in degrees, of the shape.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleHeight_scaleFactor__scaleType__scaleFrom_)|Scales the height of the shape by a specified factor.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleWidth_scaleFactor__scaleType__scaleFrom_)|Scales the width of the shape by a specified factor.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setZOrder_position_)|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[top](/javascript/api/excel/excel.shape#top)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[visible](/javascript/api/excel/excel.shape#visible)|Specifies if the shape is visible.|
||[width](/javascript/api/excel/excel.shape#width)|Specifies the width, in points, of the shape.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeId)|Gets the ID of the activated shape.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetId)|Gets the ID of the worksheet in which the shape is activated.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addGeometricShape_geometricShapeType_)|Adds a geometric shape to the worksheet.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addGroup_values_)|Groups a subset of shapes in this collection's worksheet.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addImage_base64ImageString_)|Creates an image from a base64-encoded string and adds it to the worksheet.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addLine_startLeft__startTop__endLeft__endTop__connectorType_)|Adds a line to worksheet.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addTextBox_text_)|Adds a text box to the worksheet with the provided text as the content.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getCount__)|Returns the number of shapes in the worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getItem_key_)|Gets a shape using its name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getItemAt_index_)|Gets a shape using its position in the collection.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Gets the loaded child items in this collection.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeId)|Gets the ID of the shape deactivated shape.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetId)|Gets the ID of the worksheet in which the shape is deactivated.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear__)|Clears the fill formatting of this shape.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundColor)|Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange")|
||[type](/javascript/api/excel/excel.shapefill#type)|Returns the fill type of the shape.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setSolidColor_color_)|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.shapefont#color)|HTML color code representation of the text color (e.g., "#FF0000" represents red).|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Represents the italic status of font.|
||[name](/javascript/api/excel/excel.shapefont#name)|Represents font name (e.g., "Calibri").|
||[size](/javascript/api/excel/excel.shapefont#size)|Represents font size in points (e.g., 11).|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Type of underline applied to the font.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Specifies the shape identifier.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Returns the `Shape` object associated with the group.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Returns the collection of `Shape` objects.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup__)|Ungroups any grouped shapes in the specified shape group.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashStyle)|Represents the line style of the shape.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Represents the line style of the shape.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Specifies if the line formatting of a shape element is visible.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Represents the weight of the line, in points.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subField)|Specifies the subfield that is the target property name of a rich value to sort on.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getCount__)|Gets the number of styles in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getItemAt_index_)|Gets a style based on its position in the collection.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autoFilter)|Represents the `AutoFilter` object of the table.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Gets the source of the event.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableId)|Gets the ID of the table that is added.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetId)|Gets the ID of the worksheet in which the table is added.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|Gets the information about the change detail.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onAdded)|Occurs when a new table is added in a workbook.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#onDeleted)|Occurs when the specified table is deleted in a workbook.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Gets the source of the event.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableId)|Gets the ID of the table that is deleted.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tableName)|Gets the name of the table that is deleted.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetId)|Gets the ID of the worksheet in which the table is deleted.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getCount__)|Gets the number of tables in the collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getFirst__)|Gets the first table in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItem_key_)|Gets a table by name or ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Gets the loaded child items in this collection.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autoSizeSetting)|The automatic sizing settings for the text frame.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottomMargin)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/javascript/api/excel/excel.textframe#deleteText__)|Deletes all the text in the text frame.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalAlignment)|Represents the horizontal alignment of the text frame.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontalOverflow)|Represents the horizontal overflow behavior of the text frame.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftMargin)|Represents the left margin, in points, of the text frame.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Represents the angle to which the text is oriented for the text frame.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingOrder)|Represents the reading order of the text frame, either left-to-right or right-to-left.|
||[hasText](/javascript/api/excel/excel.textframe#hasText)|Specifies if the text frame contains text.|
||[textRange](/javascript/api/excel/excel.textframe#textRange)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightMargin)|Represents the right margin, in points, of the text frame.|
||[topMargin](/javascript/api/excel/excel.textframe#topMargin)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalAlignment)|Represents the vertical alignment of the text frame.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticalOverflow)|Represents the vertical overflow behavior of the text frame.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getSubstring_start__length_)|Returns a TextRange object for the substring in the given range.|
||[font](/javascript/api/excel/excel.textrange#font)|Returns a `ShapeFont` object that represents the font attributes for the text range.|
||[text](/javascript/api/excel/excel.textrange#text)|Represents the plain text content of the text range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartDataPointTrack)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getActiveChart__)|Gets the currently active chart in the workbook.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getActiveChartOrNullObject__)|Gets the currently active chart in the workbook.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getIsActiveCollabSession__)|Returns `true` if the workbook is being edited by multiple users (through co-authoring).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getSelectedRanges__)|Gets the currently selected one or more ranges from the workbook.|
||[isDirty](/javascript/api/excel/excel.workbook#isDirty)|Specifies if changes have been made since the workbook was last saved.|
||[autoSave](/javascript/api/excel/excel.workbook#autoSave)|Specifies if the workbook is in AutoSave mode.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationEngineVersion)|Returns a number about the version of Excel Calculation Engine.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged)|Occurs when the AutoSave setting is changed on the workbook.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslySaved)|Specifies if the workbook has ever been saved locally or online.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#usePrecisionAsDisplayed)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Gets the type of the event.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enableCalculation)|Determines if Excel should recalculate the worksheet when necessary.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAll_text__criteria_)|Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAllOrNullObject_text__criteria_)|Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getRanges_address_)|Gets the `RangeAreas` object, representing one or more blocks of rectangular ranges, specified by the address or name.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autoFilter)|Represents the `AutoFilter` object of the worksheet.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalPageBreaks)|Gets the horizontal page break collection for the worksheet.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onFormatChanged)|Occurs when format changed on a specific worksheet.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pageLayout)|Gets the `PageLayout` object of the worksheet.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Returns the collection of all the Shape objects on the worksheet.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalPageBreaks)|Gets the vertical page break collection for the worksheet.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceAll_text__replacement__criteria_)|Finds and replaces the given string based on the criteria specified within the current worksheet.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Represents the information about the change detail.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onChanged)|Occurs when any worksheet in the workbook is changed.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onFormatChanged)|Occurs when any worksheet in the workbook has a format changed.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged)|Occurs when the selection changes on any worksheet.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRange_ctx_)|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRangeOrNullObject_ctx_)|Gets the range that represents the changed area of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetId)|Gets the ID of the worksheet in which the data changed.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completeMatch)|Specifies if the match needs to be complete or partial.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchCase)|Specifies if the match is case-sensitive.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
