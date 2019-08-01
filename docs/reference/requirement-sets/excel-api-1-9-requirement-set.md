---
title: Excel JavaScript API requirement set 1.9
description: 'Details about the ExcelApi 1.9 requirement set'
ms.date: 07/26/2019
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
| [Find](../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | Find values or formulas within a range or worksheet. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copy and Paste](../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | Copy values, formats, and formulas from one range to another. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](../../excel/performance.md#suspend-calculation-temporarily) | Greater control over the Excel calculation engine. | [Application](/javascript/api/excel/excel.application) |
| New Charts | Explore our new supported chart types: maps, box and whisker, waterfall, sunburst, pareto. and funnel. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | New capabilities with range formats. | [Range](/javascript/api/excel/excel.rangeformat) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.9. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.9 or earlier, see [Excel APIs in requirement set 1.9 or earlier](/javascript/api/excel?view=excel-js-1.9).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Returns the Excel calculation engine version used for the last full recalculation. Read-only.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Returns the calculation state of the application. See Excel.CalculationState for details. Read-only.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Returns the Iterative Calculation settings.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Suspends sceen updating until the next "context.sync()" is called.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Applies the AutoFilter to a range. This filters the column if column index and filter criteria are specified.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Clears the filter criteria of the AutoFilter.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Returns the Range object that represents the range to which the AutoFilter applies.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Returns the Range object that represents the range to which the AutoFilter applies.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|An array that holds all the filter criteria in the autofiltered range. Read-Only.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Indicates if the AutoFilter is enabled or not. Read-Only.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Indicates if the AutoFilter has filter criteria. Read-Only.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Applies the specified Autofilter object currently on the range.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Removes the AutoFilter for the range.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Represents the `color` property of a single border.|
||[style](/javascript/api/excel/excel.cellborder#style)|Represents the `style` property of a single border.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|Represents the `tintAndShade` property of a single border.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Represents the `weight` property of a single border.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|Represents the `format.borders.bottom` property.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|Represents the `format.borders.diagonalDown` property.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|Represents the `format.borders.diagonalUp` property.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Represents the `format.borders.horizontal` property.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Represents the `format.borders.left` property.|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|Represents the `format.borders.right` property.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Represents the `format.borders.top` property.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Represents the `format.borders.vertical` property.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|Represents the `addressLocal` property.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Represents the `hidden` property.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Represents the `format.fill.color` property.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Represents the `format.fill.pattern` property.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Represents the `format.fill.patternColor` property.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Represents the `format.fill.patternTintAndShade` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Represents the `format.fill.tintAndShade` property.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Represents the `format.font.bold` property.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Represents the `format.font.color` property.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Represents the `format.font.italic` property.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Represents the `format.font.name` property.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Represents the `format.font.size` property.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Represents the `format.font.strikethrough` property.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Represents the `format.font.subscript` property.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Represents the `format.font.superscript` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|Represents the `format.font.tintAndShade` property.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Represents the `format.font.underline` property.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|Represents the `autoIndent` property.|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Represents the `borders` property.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Represents the `fill` property.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|Represents the `font` property.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|Represents the `horizontalAlignment` property.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|Represents the `indentLevel` property.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Represents the `protection` property.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|Represents the `readingOrder` property.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|Represents the `shrinkToFit` property.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|Represents the `textOrientation` property.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|Represents the `useStandardHeight` property.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|Represents the `useStandardWidth` property.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|Represents the `verticalAlignment` property.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Represents the `wrapText` property.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|Represents the `format.protection.formulaHidden` property.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Represents the `format.protection.locked` property.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Represents the value after changed. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Represents the value before changed. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Represents the type of value after changed|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Represents the type of value before changed|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Activates the chart in the Excel UI.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Encapsulates the options for a pivot chart. Read-only.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Returns or sets color scheme of the chart. Read/Write.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|Specifies whether or not chart area of the chart has rounded corners. Read/Write.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Specifies whether or not the bin overflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Specifies whether or not the bin underflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Returns or sets the bin count of a histogram chart or pareto chart. Read/Write.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Returns or sets the bin overflow value of a histogram chart or pareto chart. Read/Write.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Returns or sets the bin's type for a histogram chart or pareto chart. Read/Write.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Returns or sets the bin underflow value of a histogram chart or pareto chart. Read/Write.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Returns or sets the bin width value of a histogram chart or pareto chart. Read/Write.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Returns or sets the quartile calculation type of a box and whisker chart. Read/Write.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Specifies whether or not the inner points are shown in a box and whisker chart. Read/Write.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Specifies whether or not the mean line is shown in a box and whisker chart. Read/Write.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Specifies whether or not the mean marker is shown in a box and whisker chart. Read/Write.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Specifies whether or not outlier points are shown in a box and whisker chart. Read/Write.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Specifies whether or not the error bars have an end style cap.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Specifies which parts of the error bars to include.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Specifies the formatting type of the error bars.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|The type of range marked by the error bars.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Specifies whether or not the error bars are displayed.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Represents the chart line formatting.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Returns or sets the series map labels strategy of a region map chart. Read/Write.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Returns or sets the series mapping level of a region map chart. Read/Write.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Returns or sets the series projection type of a region map chart. Read/Write.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Specifies whether or not to display the axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Specifies whether or not to display the legend field buttons on a PivotChart|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Specifies whether or not to display the report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Specifies whether or not to display the show value field buttons on a PivotChart|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts. Read/Write.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Returns or sets the color for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Returns or sets the type for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Returns or sets the maximum value of a region map chart series. Read/Write.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Returns or sets the color for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Returns or sets the type for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Returns or sets the midpoint value of a region map chart series. Read/Write.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Returns or sets the color for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Returns or sets the type for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Returns or sets the minimum value of a region map chart series. Read/Write.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Returns or sets series gradient style of a region map chart. Read/Write.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Returns or sets the fill color for negative data points in a series. Read/Write.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Returns or sets the series parent label strategy area for a treemap chart. Read/Write.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Encapsulates the bin options for histogram charts and pareto charts. Read-only.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Encapsulates the options for the box and whisker charts. Read-only.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Encapsulates the options for a region map chart. Read-only.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Represents the error bar object of a chart series.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Represents the error bar object of a chart series.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Specifies whether or not connector lines are shown in waterfall charts. Read/Write.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Specifies whether or not leader lines are displayed for each data label in the series. Read/Write.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Returns or sets the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/Write.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Represents the `addressLocal` property.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Represents the `columnIndex` property.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Returns the RangeAreas, comprising one or more rectangular ranges, the conditonal format is applied to. Read-only.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will throw an ItemNotFound error.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will return null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|The property used by the filter to do rich filter on richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Returns the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Returns the Shape object for the geometric shape. Read-only.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Returns the number of shapes in the shape group. Read-only.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Gets a shape using its Name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Gets a shape based on its position in the collection.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Gets the loaded child items in this collection.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Gets or sets the center footer of the worksheet.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Gets or sets the center header of the worksheet.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Gets or sets the left footer of the worksheet.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Gets or sets the left header of the worksheet.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Gets or sets the right footer of the worksheet.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Gets or sets the right header of the worksheet.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|The general header/footer, used for all pages unless even/odd or first page is specified.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|The first page header/footer, for all other pages general or even/odd is used.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|The header/footer to use for odd pages, even header/footer needs to be specified for even pages.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Gets or sets the state of which headers/footers are set. See Excel.HeaderFooterState for details.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Returns the format of the image. Read-only.|
||[id](/javascript/api/excel/excel.image#id)|Represents the shape identifier for the image object. Read-only.|
||[shape](/javascript/api/excel/excel.image#shape)|Returns the Shape object associated with the image. Read-only.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|Represents the length of the arrowhead at the beginning of the specified line.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|Represents the style of the arrowhead at the beginning of the specified line.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|Represents the width of the arrowhead at the beginning of the specified line.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|Attaches the beginning of the specified connector to a specified shape.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|Attaches the end of the specified connector to a specified shape.|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|Represents the connector type for the line.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|Detaches the beginning of the specified connector from a shape.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|Detaches the end of the specified connector from a shape.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|Represents the length of the arrowhead at the end of the specified line.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|Represents the style of the arrowhead at the end of the specified line.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|Represents the width of the arrowhead at the end of the specified line.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Represents the shape to which the beginning of the specified line is attached. Read-only.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Represents the connection site to which the beginning of a connector is connected. Read-only. Returns null when the beginning of the line is not attached to any shape.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Represents the shape to which the end of the specified line is attached. Read-only.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Represents the connection site to which the end of a connector is connected. Read-only. Returns null when the end of the line is not attached to any shape.|
||[id](/javascript/api/excel/excel.line#id)|Represents the shape identifier. Read-only.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Specifies whether or not the beginning of the specified line is connected to a shape. Read-only.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Specifies whether or not the end of the specified line is connected to a shape. Read-only.|
||[shape](/javascript/api/excel/excel.line#shape)|Returns the Shape object associated with the line. Read-only.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Deletes a page break object.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Gets the first cell after the page break.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Represents the column index for the page break|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Represents the row index for the page break|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Adds a page break before the top-left cell of the range specified.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Gets the number of page breaks in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Gets a page break object via the index.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Gets the loaded child items in this collection.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Resets all manual page breaks in the collection.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|Gets or sets the worksheet's black and white print option.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|Gets or sets the worksheet's bottom page margin to use for printing in points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|Gets or sets the worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|Gets or sets the worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|Gets or sets the worksheet's draft mode option. If true the sheet will be printed without graphics.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|Gets or sets the worksheet's first page number to print. Null value represents "auto" page numbering.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|Gets or sets the worksheet's footer margin, in points, for use when printing.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, an ItemNotFound error will be thrown.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, a null object will be returned.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Gets the range object representing the title columns.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Gets the range object representing the title columns. If not set, this will return a null object.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Gets the range object representing the title rows.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Gets the range object representing the title rows. If not set, this will return a null object.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|Gets or sets the worksheet's header margin, in points, for use when printing.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|Gets or sets the worksheet's left margin, in points, for use when printing.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|Gets or sets the worksheet's orientation of the page.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|Gets or sets the worksheet's paper size of the page.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Gets or sets whether the worksheet's comments should be displayed when printing.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|Gets or sets the worksheet's print errors option.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Gets or sets the worksheet's print gridlines flag. This flag determines whether gridlines will be printed or not.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Gets or sets the worksheet's print headings flag. This flag determines whether headings will be printed or not.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|Gets or sets the worksheet's page print order option. This specifies the order to use for processing the page number printed.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Header and footer configuration for the worksheet.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|Gets or sets the worksheet's right margin, in points, for use when printing.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Sets the worksheet's print area.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Sets the worksheet's page margins with units.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Gets or sets the worksheet's top margin, in points, for use when printing.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Gets or sets the worksheet's print zoom options.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Represents the page layout bottom margin in the unit specified to use for printing.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Represents the page layout footer margin in the unit specified to use for printing.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Represents the page layout header margin in the unit specified to use for printing.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Represents the page layout left margin in the unit specified to use for printing.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Represents the page layout right margin in the unit specified to use for printing.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Represents the page layout top margin in the unit specified to use for printing.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Number of pages to fit horizontally. This value can be null if percentage scale is used.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|Print page scale value can be between 10 and 400. This value can be null if fit to page tall or wide is specified.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Number of pages to fit vertically. This value can be null if percentage scale is used.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Specifies whether formatting will be automatically formatted when it’s refreshed or when fields are moved|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Specifies whether formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context. This behaves identically to applying an autosort from the UI.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Specifies whether the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Specifies whether the PivotTable uses custom lists when sorting.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Fills range from the current range to the destination range.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Converts the range cells with datatypes into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Converts the range cells into linked datatype in the worksheet.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copies cell data or formatting from the source range or RangeAreas to the current range.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Finds the given string based on the criteria specified.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Finds the given string based on the criteria specified.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Does FlashFill to current range.Flash Fill will automatically fills data when it senses a pattern, so the range must be single column range and have data around in order to find pattern.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Returns a 2D array, encapsulating the data for each cell's font, fill, borders, alignment, and other properties.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Returns a single-dimensional array, encapsulating the data for each column's font, fill, borders, alignment, and other properties.  For properties that are not consistent across each cell within a given column, null will be returned.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Returns a single-dimensional array, encapsulating the data for each row's font, fill, borders, alignment, and other properties.  For properties that are not consistent across each cell within a given row, null will be returned.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Gets a scoped collection of tables that overlap with the range.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Represents the data type state of each cell. Read-only.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Removes duplicate values from the range specified by the columns.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Finds and replaces the given string based on the criteria specified within the current range.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Updates the range based on a 2D array of cell properties , encapsulating things like font, fill, borders, alignment, and so forth.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Updates the range based on a single-dimensional array of column properties, encapsulating things like font, fill, borders, alignment, and so forth.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Set a range to be recalculated when the next recalculation occurs.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Updates the range based on a single-dimensional array of row properties, encapsulating things like font, fill, borders, alignment, and so forth.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Calculates all cells in the RangeAreas.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Converts all cells in the RangeAreas with datatypes into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Converts all cells in the RangeAreas into linked datatype.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Returns a RangeAreas object that represents the entire columns of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11, H2", it returns a RangeAreas that represents columns "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Returns a RangeAreas object that represents the entire rows of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11", it returns a RangeAreas that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, an ItemNotFound error will be thrown.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, a null object is returned.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Returns an RangeAreas object that is shifted by the specific row and column offset. The dimension of the returned RangeAreas will match the original object. If the resulting RangeAreas is forced outside the bounds of the worksheet grid, an error will be thrown.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Returns a scoped collection of tables that overlap with any range in this RangeAreas object.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Returns the RageAreas reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g. "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Returns the RageAreas reference in the user locale. Read-only.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Returns the number of rectangular ranges that comprise this RangeAreas object.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Returns a collection of rectangular ranges that comprise this RangeAreas object.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Returns the number of cells in the RangeAreas object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Returns a collection of ConditionalFormats that intersect with any cells in this RangeAreas object. Read-only.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Returns a dataValidation object for all ranges in the RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Returns a rangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object. Read-only.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Indicates whether all the ranges on this RangeAreas object represent entire columns (e.g., "A:C, Q:Z"). Read-only.|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Indicates whether all the ranges on this RangeAreas object represent entire rows (e.g., "1:3, 5:7"). Read-only.|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Returns the worksheet for the current RangeAreas. Read-only.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Sets the RangeAreas to be recalculated when the next recalculation occurs.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Represents the style for all ranges in this RangeAreas object.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Borders, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Returns the number of ranges in the RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Returns the range object based on its position in the RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Gets the loaded child items in this collection.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Gets or sets the pattern of a Range. See Excel.FillPattern for details. LinearGradient and RectangularGradient are not supported.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Sets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Returns or sets a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Represents the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Represents the Subscript status of font.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Represents the Superscript status of font.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Indicates if text is automatically indented when text alignment is set to equal distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|An integer from 0 to 250 that indicates the indent level.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|The reading order for the range.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Number of duplicated rows removed by the operation.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Number of remaining unique rows present in the resulting range.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Specifies whether the match needs to be complete or partial. Default is false (partial).|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Specifies whether the match is case sensitive. Default is false (insensitive).|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Represents the `addressLocal` property.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Represents the `rowIndex` property.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Specifies whether the match needs to be complete or partial. A complete match matches the entire contents of the cell. Default is false (partial).|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Specifies whether the match is case sensitive. Default is false (insensitive).|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Specifies the search direction. Default is forward. See Excel.SearchDirection.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Represents the `format` property.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Represents the `hyperlink` property.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Represents the `style` property.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Represents the `columnHidden` property.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[format: Excel.CellPropertiesFormat & {
            columnWidth?](/javascript/api/excel/excel.settablecolumnproperties#format)|Represents the `format` property.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel.CellPropertiesFormat & {
            rowHeight?](/javascript/api/excel/excel.settablerowproperties#format)|Represents the `format` property.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Represents the `rowHidden` property.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Returns or sets the alternative description text for a Shape object.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Returns or sets the alternative title text for a Shape object.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Removes the shape from the worksheet.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Represents the geometric shape type of this geometric shape. See Excel.GeometricShapeType for details. Returns null if the shape type is not "GeometricShape".|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Converts the shape to an image and returns the image as a base64-encoded string. The DPI is 96. The only supported formats are `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, and `Excel.PictureFormat.GIF`.|
||[height](/javascript/api/excel/excel.shape#height)|Represents the height, in points, of the shape.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Moves the shape horizontally by the specified number of points.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|Rotates the shape clockwise around the z-axis by the specified number of degrees.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Moves the shape vertically by the specified number of points.|
||[left](/javascript/api/excel/excel.shape#left)|The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Specifies whether or not the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.shape#name)|Represents the name of the shape.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Returns the number of connection sites on this shape. Read-only.|
||[fill](/javascript/api/excel/excel.shape#fill)|Returns the fill formatting of this shape. Read-only.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Returns the geometric shape associated with the shape. An error will be thrown if the shape type is not "GeometricShape".|
||[group](/javascript/api/excel/excel.shape#group)|Returns the shape group associated with the shape. An error will be thrown if the shape type is not "GroupShape".|
||[id](/javascript/api/excel/excel.shape#id)|Represents the shape identifier. Read-only.|
||[image](/javascript/api/excel/excel.shape#image)|Returns the image associated with the shape. An error will be thrown if the shape type is not "Image".|
||[level](/javascript/api/excel/excel.shape#level)|Represents the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.|
||[line](/javascript/api/excel/excel.shape#line)|Returns the line associated with the shape. An error will be thrown if the shape type is not "Line".|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Returns the line formatting of this shape. Read-only.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Occurs when the shape is activated.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Occurs when the shape is deactivated.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Represents the parent group of this shape.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Returns the text frame object of this shape. Read only.|
||[type](/javascript/api/excel/excel.shape#type)|Returns the type of this shape. See Excel.ShapeType for details. Read-only.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack. Read-only.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Represents the rotation, in degrees, of the shape.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[top](/javascript/api/excel/excel.shape#top)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[visible](/javascript/api/excel/excel.shape#visible)|Represents the visibility of this shape.|
||[width](/javascript/api/excel/excel.shape#width)|Represents the width, in points, of the shape.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Gets the id of the activated shape.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Gets the id of the worksheet in which the shape is activated.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Adds a geometric shape to the worksheet. Returns a Shape object that represents the new shape.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Groups a subset of shapes in this collection's worksheet. Returns a Shape object that represents the new group of shapes.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Creates an image from a base64-encoded string and adds it to the worksheet. Returns the Shape object that represents the new image.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adds a line to worksheet. Returns a Shape object that represents the new line.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Adds a text box to the worksheet with the provided text as the content. Returns a Shape object that represents the new text box.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Returns the number of shapes in the worksheet. Read-only.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Gets a shape using its Name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Gets a shape using its position in the collection.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Gets the loaded child items in this collection.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Gets the id of the shape deactivated shape.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Gets the id of the worksheet in which the shape is deactivated.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Clears the fill formatting of this shape.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Represents the shape fill foreground color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
||[type](/javascript/api/excel/excel.shapefill#type)|Returns the fill type of the shape. Read-only. See Excel.ShapeFillType for details.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Sets the fill formatting of the shape to a uniform color. This changes the fill type to "Solid".|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Returns or sets the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|
||[color](/javascript/api/excel/excel.shapefont#color)|The HTML color code representation of the text color (e.g. "#FF0000" represents red). Returns null if the TextRange includes text fragments with different colors.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Represents the italic status of font. Returns null if the TextRange includes both italic and non-italic text fragments.|
||[name](/javascript/api/excel/excel.shapefont#name)|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.|
||[size](/javascript/api/excel/excel.shapefont#size)|Represents font size in points (e.g. 11). Returns null if the TextRange includes text fragments with different font sizes.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Type of underline applied to the font. Returns null if the TextRange includes text fragments with different underline styles. See Excel.ShapeFontUnderlineStyle for details.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Represents the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Returns the Shape object associated with the group. Read-only.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Returns the collection of Shape objects. Read-only.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Ungroups any grouped shapes in the specified shape group.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent dash styles. See Excel.ShapeLineStyle for details.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See Excel.ShapeLineStyle for details.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Represents whether or not the line formatting of a shape element is visible. Returns null when the shape has inconsistent visibilities.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Represents the weight of the line, in points. Returns null when the line is not visible or there are inconsistent line weights.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Represents the subfield that is the target property name of a rich value to sort on.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Gets the number of styles in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Gets a style based on its position in the collection.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|Represents the AutoFilter object of the table. Read-Only.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Gets the id of the table that is added.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Gets the id of the worksheet in which the table is added.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|Represents the information about the change detail|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Occurs when new table is added in a workbook.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Occurs when the specified table is deleted in a workbook.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Specifies the source of the event. See Excel.EventSource for details.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Specifies the id of the table that is deleted.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Specifies the name of the table that is deleted.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Specifies the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Specifies the id of the worksheet in which the table is deleted.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Gets the number of tables in the collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Gets the first table in the collection. The tables in the collection are sorted top to bottom and left to right, such that top left table is the first table in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Gets a table by Name or ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Gets the loaded child items in this collection.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|Gets or sets the automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Deletes all the text in the text frame.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Represents the horizontal alignment of the text frame. See Excel.ShapeTextHorizontalAlignment for details.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Represents the horizontal overflow behavior of the text frame. See Excel.ShapeTextHorizontalOverflow for details.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Represents the left margin, in points, of the text frame.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Represents the text orientation of the text frame. See Excel.ShapeTextOrientation for details.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Represents the reading order of the text frame, either left-to-right or right-to-left. See Excel.ShapeTextReadingOrder for details.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Specifies whether the text frame contains text.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See Excel.TextRange for details.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Represents the right margin, in points, of the text frame.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Represents the vertical alignment of the text frame. See Excel.ShapeTextVerticalAlignment for details.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Represents the vertical overflow behavior of the text frame. See Excel.ShapeTextVerticalOverflow for details.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Returns a TextRange object for the substring in the given range.|
||[font](/javascript/api/excel/excel.textrange#font)|Returns a ShapeFont object that represents the font attributes for the text range. Read-only.|
||[text](/javascript/api/excel/excel.textrange#text)|Represents the plain text content of the text range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Gets the currently active chart in the workbook. If there is no active chart, will throw exception when invoke this statement|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Gets the currently active chart in the workbook. If there is no active chart, will return null object|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|True if the workbook is being edited by multiple users (co-authoring).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Gets the currently selected one or more ranges from the workbook. Unlike getSelectedRange(), this method returns a RangeAreas object that represents all the selected ranges.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Specifies whether or not changes have been made since the workbook was last saved.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Specifies whether or not the workbook is in autosave mode. Read-Only.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Returns a number about the version of Excel Calculation Engine. Read-Only.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Occurs when the autoSave setting is changed on the workbook.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Specifies whether or not the workbook has ever been saved locally or online. Read-Only.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Gets or sets the enableCalculation property of the worksheet.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Gets the RangeAreas object, representing one or more blocks of rectangular ranges, specified by the address or name.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Represents the AutoFilter object of the worksheet. Read-Only.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Occurs when format changed on a specific worksheet.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Gets the PageLayout object of the worksheet.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Returns the collection of all the Shape objects on the worksheet. Read-only.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Finds and replaces the given string based on the criteria specified within the current worksheet.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Represents the information about the change detail|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Occurs when any worksheet in the workbook is changed.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Occurs when any worksheet in the workbook has format changed.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Occurs when the selection changes on any worksheet.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Gets the range that represents the changed area of a specific worksheet. It might return null object.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Specifies whether the match needs to be complete or partial. A complete match matches the entire contents of the cell. Default is false (partial).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Specifies whether the match is case sensitive. Default is false (insensitive).|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.9)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
