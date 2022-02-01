---
title: Excel JavaScript API requirement set 1.9
description: 'Details about the ExcelApi 1.9 requirement set.'
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
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
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#excel-excel-application-calculationengineversion-member)|Returns the Excel calculation engine version used for the last full recalculation.|
||[calculationState](/javascript/api/excel/excel.application#excel-excel-application-calculationstate-member)|Returns the calculation state of the application.|
||[iterativeCalculation](/javascript/api/excel/excel.application#excel-excel-application-iterativecalculation-member)|Returns the iterative calculation settings.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendscreenupdatinguntilnextsync-member(1))|Suspends screen updating until the next `context.sync()` is called.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-apply-member(1))|Applies the AutoFilter to a range.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcriteria-member(1))|Clears the filter criteria and sort state of the AutoFilter.|
||[criteria](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-criteria-member)|An array that holds all the filter criteria in the autofiltered range.|
||[enabled](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-enabled-member)|Specifies if the AutoFilter is enabled.|
||[getRange()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrange-member(1))|Returns the `Range` object that represents the range to which the AutoFilter applies.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrangeornullobject-member(1))|Returns the `Range` object that represents the range to which the AutoFilter applies.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-isdatafiltered-member)|Specifies if the AutoFilter has filter criteria.|
||[reapply()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-reapply-member(1))|Applies the specified Autofilter object currently on the range.|
||[remove()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-remove-member(1))|Removes the AutoFilter for the range.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-color-member)|Represents the `color` property of a single border.|
||[style](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-style-member)|Represents the `style` property of a single border.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-tintandshade-member)|Represents the `tintAndShade` property of a single border.|
||[weight](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-weight-member)|Represents the `weight` property of a single border.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-bottom-member)|Represents the `format.borders.bottom` property.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonaldown-member)|Represents the `format.borders.diagonalDown` property.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonalup-member)|Represents the `format.borders.diagonalUp` property.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-horizontal-member)|Represents the `format.borders.horizontal` property.|
||[left](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-left-member)|Represents the `format.borders.left` property.|
||[right](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-right-member)|Represents the `format.borders.right` property.|
||[top](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-top-member)|Represents the `format.borders.top` property.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-vertical-member)|Represents the `format.borders.vertical` property.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-address-member)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-addresslocal-member)|Represents the `addressLocal` property.|
||[hidden](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-hidden-member)|Represents the `hidden` property.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-color-member)|Represents the `format.fill.color` property.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-pattern-member)|Represents the `format.fill.pattern` property.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterncolor-member)|Represents the `format.fill.patternColor` property.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterntintandshade-member)|Represents the `format.fill.patternTintAndShade` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-tintandshade-member)|Represents the `format.fill.tintAndShade` property.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-bold-member)|Represents the `format.font.bold` property.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-color-member)|Represents the `format.font.color` property.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-italic-member)|Represents the `format.font.italic` property.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-name-member)|Represents the `format.font.name` property.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-size-member)|Represents the `format.font.size` property.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-strikethrough-member)|Represents the `format.font.strikethrough` property.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-subscript-member)|Represents the `format.font.subscript` property.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-superscript-member)|Represents the `format.font.superscript` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-tintandshade-member)|Represents the `format.font.tintAndShade` property.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-underline-member)|Represents the `format.font.underline` property.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-autoindent-member)|Represents the `autoIndent` property.|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-borders-member)|Represents the `borders` property.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-fill-member)|Represents the `fill` property.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-font-member)|Represents the `font` property.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-horizontalalignment-member)|Represents the `horizontalAlignment` property.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-indentlevel-member)|Represents the `indentLevel` property.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-protection-member)|Represents the `protection` property.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-readingorder-member)|Represents the `readingOrder` property.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-shrinktofit-member)|Represents the `shrinkToFit` property.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-textorientation-member)|Represents the `textOrientation` property.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member)|Represents the `useStandardHeight` property.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member)|Represents the `useStandardWidth` property.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-verticalalignment-member)|Represents the `verticalAlignment` property.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-wraptext-member)|Represents the `wrapText` property.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-formulahidden-member)|Represents the `format.protection.formulaHidden` property.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-locked-member)|Represents the `format.protection.locked` property.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valueafter-member)|Represents the value after the change.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuebefore-member)|Represents the value before the change.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypeafter-member)|Represents the type of value after the change.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypebefore-member)|Represents the type of value before the change.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#excel-excel-chart-activate-member(1))|Activates the chart in the Excel UI.|
||[pivotOptions](/javascript/api/excel/excel.chart#excel-excel-chart-pivotoptions-member)|Encapsulates the options for a pivot chart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-colorscheme-member)|Specifies the color scheme of the chart.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-roundedcorners-member)|Specifies if the chart area of the chart has rounded corners.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-linknumberformat-member)|Specifies if the number format is linked to the cells.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowoverflow-member)|Specifies if bin overflow is enabled in a histogram chart or pareto chart.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowunderflow-member)|Specifies if bin underflow is enabled in a histogram chart or pareto chart.|
||[count](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-count-member)|Specifies the bin count of a histogram chart or pareto chart.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-overflowvalue-member)|Specifies the bin overflow value of a histogram chart or pareto chart.|
||[type](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-type-member)|Specifies the bin's type for a histogram chart or pareto chart.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-underflowvalue-member)|Specifies the bin underflow value of a histogram chart or pareto chart.|
||[width](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-width-member)|Specifies the bin width value of a histogram chart or pareto chart.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-quartilecalculation-member)|Specifies if the quartile calculation type of a box and whisker chart.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showinnerpoints-member)|Specifies if inner points are shown in a box and whisker chart.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanline-member)|Specifies if the mean line is shown in a box and whisker chart.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanmarker-member)|Specifies if the mean marker is shown in a box and whisker chart.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showoutlierpoints-member)|Specifies if outlier points are shown in a box and whisker chart.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-linknumberformat-member)|Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-linknumberformat-member)|Specifies if the number format is linked to the cells.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-endstylecap-member)|Specifies if error bars have an end style cap.|
||[format](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-format-member)|Specifies the formatting type of the error bars.|
||[include](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-include-member)|Specifies which parts of the error bars to include.|
||[type](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-type-member)|The type of range marked by the error bars.|
||[visible](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-visible-member)|Specifies whether the error bars are displayed.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#excel-excel-charterrorbarsformat-line-member)|Represents the chart line formatting.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-labelstrategy-member)|Specifies the series map labels strategy of a region map chart.|
||[level](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-level-member)|Specifies the series mapping level of a region map chart.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-projectiontype-member)|Specifies the series projection type of a region map chart.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showaxisfieldbuttons-member)|Specifies whether to display the axis field buttons on a PivotChart.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showlegendfieldbuttons-member)|Specifies whether to display the legend field buttons on a PivotChart.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showreportfilterfieldbuttons-member)|Specifies whether to display the report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showvaluefieldbuttons-member)|Specifies whether to display the show value field buttons on a PivotChart.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[binOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-binoptions-member)|Encapsulates the bin options for histogram charts and pareto charts.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-boxwhiskeroptions-member)|Encapsulates the options for the box and whisker charts.|
||[bubbleScale](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-bubblescale-member)|This can be an integer value from 0 (zero) to 300, representing the percentage of the default size.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumcolor-member)|Specifies the color for maximum value of a region map chart series.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumtype-member)|Specifies the type for maximum value of a region map chart series.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumvalue-member)|Specifies the maximum value of a region map chart series.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointcolor-member)|Specifies the color for the midpoint value of a region map chart series.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointtype-member)|Specifies the type for the midpoint value of a region map chart series.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointvalue-member)|Specifies the midpoint value of a region map chart series.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumcolor-member)|Specifies the color for the minimum value of a region map chart series.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumtype-member)|Specifies the type for the minimum value of a region map chart series.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumvalue-member)|Specifies the minimum value of a region map chart series.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientstyle-member)|Specifies the series gradient style of a region map chart.|
||[invertColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertcolor-member)|Specifies the fill color for negative data points in a series.|
||[mapOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-mapoptions-member)|Encapsulates the options for a region map chart.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-parentlabelstrategy-member)|Specifies the series parent label strategy area for a treemap chart.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showconnectorlines-member)|Specifies whether connector lines are shown in waterfall charts.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showleaderlines-member)|Specifies whether leader lines are displayed for each data label in the series.|
||[splitValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splitvalue-member)|Specifies the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-xerrorbars-member)|Represents the error bar object of a chart series.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-yerrorbars-member)|Represents the error bar object of a chart series.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-linknumberformat-member)|Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-address-member)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-addresslocal-member)|Represents the `addressLocal` property.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-columnindex-member)|Represents the `columnIndex` property.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getranges-member(1))|Returns the `RangeAreas`, comprising one or more rectangular ranges, to which the conditonal format is applied.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcells-member(1))|Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcellsornullobject-member(1))|Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#excel-excel-filtercriteria-subfield-member)|The property used by the filter to do a rich filter on rich values.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-id-member)|Returns the shape identifier.|
||[shape](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-shape-member)|Returns the `Shape` object for the geometric shape.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getcount-member(1))|Returns the number of shapes in the shape group.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitem-member(1))|Gets a shape using its name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemat-member(1))|Gets a shape based on its position in the collection.|
||[items](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-items-member)|Gets the loaded child items in this collection.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerfooter-member)|The center footer of the worksheet.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerheader-member)|The center header of the worksheet.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftfooter-member)|The left footer of the worksheet.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftheader-member)|The left header of the worksheet.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightfooter-member)|The right footer of the worksheet.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightheader-member)|The right header of the worksheet.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-defaultforallpages-member)|The general header/footer, used for all pages unless even/odd or first page is specified.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-evenpages-member)|The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-firstpage-member)|The first page header/footer, for all other pages general or even/odd is used.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-oddpages-member)|The header/footer to use for odd pages, even header/footer needs to be specified for even pages.|
||[state](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-state-member)|The state by which headers/footers are set.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetmargins-member)|Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetscale-member)|Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#excel-excel-image-format-member)|Returns the format of the image.|
||[id](/javascript/api/excel/excel.image#excel-excel-image-id-member)|Specifies the shape identifier for the image object.|
||[shape](/javascript/api/excel/excel.image#excel-excel-image-shape-member)|Returns the `Shape` object associated with the image.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-enabled-member)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxchange-member)|Specifies the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxiteration-member)|Specifies the maximum number of iterations that Excel can use to resolve a circular reference.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadlength-member)|Represents the length of the arrowhead at the beginning of the specified line.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadstyle-member)|Represents the style of the arrowhead at the beginning of the specified line.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadwidth-member)|Represents the width of the arrowhead at the beginning of the specified line.|
||[beginConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedshape-member)|Represents the shape to which the beginning of the specified line is attached.|
||[beginConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedsite-member)|Represents the connection site to which the beginning of a connector is connected.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectbeginshape-member(1))|Attaches the beginning of the specified connector to a specified shape.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectendshape-member(1))|Attaches the end of the specified connector to a specified shape.|
||[connectorType](/javascript/api/excel/excel.line#excel-excel-line-connectortype-member)|Represents the connector type for the line.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectbeginshape-member(1))|Detaches the beginning of the specified connector from a shape.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectendshape-member(1))|Detaches the end of the specified connector from a shape.|
||[endArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadlength-member)|Represents the length of the arrowhead at the end of the specified line.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadstyle-member)|Represents the style of the arrowhead at the end of the specified line.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadwidth-member)|Represents the width of the arrowhead at the end of the specified line.|
||[endConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-endconnectedshape-member)|Represents the shape to which the end of the specified line is attached.|
||[endConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-endconnectedsite-member)|Represents the connection site to which the end of a connector is connected.|
||[id](/javascript/api/excel/excel.line#excel-excel-line-id-member)|Specifies the shape identifier.|
||[isBeginConnected](/javascript/api/excel/excel.line#excel-excel-line-isbeginconnected-member)|Specifies if the beginning of the specified line is connected to a shape.|
||[isEndConnected](/javascript/api/excel/excel.line#excel-excel-line-isendconnected-member)|Specifies if the end of the specified line is connected to a shape.|
||[shape](/javascript/api/excel/excel.line#excel-excel-line-shape-member)|Returns the `Shape` object associated with the line.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[columnIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-columnindex-member)|Specifies the column index for the page break.|
||[delete()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-delete-member(1))|Deletes a page break object.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-getcellafterbreak-member(1))|Gets the first cell after the page break.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-rowindex-member)|Specifies the row index for the page break.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-add-member(1))|Adds a page break before the top-left cell of the range specified.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getcount-member(1))|Gets the number of page breaks in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getitem-member(1))|Gets a page break object via the index.|
||[items](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-items-member)|Gets the loaded child items in this collection.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-removepagebreaks-member(1))|Resets all manual page breaks in the collection.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-blackandwhite-member)|The worksheet's black and white print option.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-bottommargin-member)|The worksheet's bottom page margin to use for printing in points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centerhorizontally-member)|The worksheet's center horizontally flag.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centervertically-member)|The worksheet's center vertically flag.|
||[draftMode](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-draftmode-member)|The worksheet's draft mode option.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-firstpagenumber-member)|The worksheet's first page number to print.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-footermargin-member)|The worksheet's footer margin, in points, for use when printing.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintarea-member(1))|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintareaornullobject-member(1))|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumns-member(1))|Gets the range object representing the title columns.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumnsornullobject-member(1))|Gets the range object representing the title columns.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerows-member(1))|Gets the range object representing the title rows.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerowsornullobject-member(1))|Gets the range object representing the title rows.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headermargin-member)|The worksheet's header margin, in points, for use when printing.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headersfooters-member)|Header and footer configuration for the worksheet.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-leftmargin-member)|The worksheet's left margin, in points, for use when printing.|
||[orientation](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-orientation-member)|The worksheet's orientation of the page.|
||[paperSize](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-papersize-member)|The worksheet's paper size of the page.|
||[printComments](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printcomments-member)|Specifies if the worksheet's comments should be displayed when printing.|
||[printErrors](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printerrors-member)|The worksheet's print errors option.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printgridlines-member)|Specifies if the worksheet's gridlines will be printed.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printheadings-member)|Specifies if the worksheet's headings will be printed.|
||[printOrder](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printorder-member)|The worksheet's page print order option.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-rightmargin-member)|The worksheet's right margin, in points, for use when printing.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintarea-member(1))|Sets the worksheet's print area.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintmargins-member(1))|Sets the worksheet's page margins with units.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlecolumns-member(1))|Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlerows-member(1))|Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.|
||[topMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-topmargin-member)|The worksheet's top margin, in points, for use when printing.|
||[zoom](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-zoom-member)|The worksheet's print zoom options.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-bottom-member)|Specifies the page layout bottom margin in the unit specified to use for printing.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-footer-member)|Specifies the page layout footer margin in the unit specified to use for printing.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-header-member)|Specifies the page layout header margin in the unit specified to use for printing.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-left-member)|Specifies the page layout left margin in the unit specified to use for printing.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-right-member)|Specifies the page layout right margin in the unit specified to use for printing.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-top-member)|Specifies the page layout top margin in the unit specified to use for printing.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-horizontalfittopages-member)|Number of pages to fit horizontally.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-scale-member)|Print page scale value can be between 10 and 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-verticalfittopages-member)|Number of pages to fit vertically.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbyvalues-member(1))|Sorts the PivotField by specified values in a given scope.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-autoformat-member)|Specifies if formatting will be automatically formatted when it’s refreshed or when fields are moved.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatahierarchy-member(1))|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getpivotitems-member(1))|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-preserveformatting-member)|Specifies if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setautosortoncell-member(1))|Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-enabledatavalueediting-member)|Specifies if the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-usecustomsortlists-member)|Specifies if the PivotTable uses custom lists when sorting.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#excel-excel-range-autofill-member(1))|Fills range from the current range to the destination range using the specified AutoFill logic.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#excel-excel-range-convertdatatypetotext-member(1))|Converts the range cells with data types into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#excel-excel-range-converttolinkeddatatype-member(1))|Converts the range cells into linked data types in the worksheet.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1))|Copies cell data or formatting from the source range or `RangeAreas` to the current range.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-find-member(1))|Finds the given string based on the criteria specified.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-findornullobject-member(1))|Finds the given string based on the criteria specified.|
||[flashFill()](/javascript/api/excel/excel.range#excel-excel-range-flashfill-member(1))|Does a Flash Fill to the current range.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcellproperties-member(1))|Returns a 2D array, encapsulating the data for each cell's font, fill, borders, alignment, and other properties.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcolumnproperties-member(1))|Returns a single-dimensional array, encapsulating the data for each column's font, fill, borders, alignment, and other properties.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getrowproperties-member(1))|Returns a single-dimensional array, encapsulating the data for each row's font, fill, borders, alignment, and other properties.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1))|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1))|Gets the `RangeAreas` object, comprising one or more ranges, that represents all the cells that match the specified type and value.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-gettables-member(1))|Gets a scoped collection of tables that overlap with the range.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#excel-excel-range-linkeddatatypestate-member)|Represents the data type state of each cell.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))|Removes duplicate values from the range specified by the columns.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#excel-excel-range-replaceall-member(1))|Finds and replaces the given string based on the criteria specified within the current range.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#excel-excel-range-setcellproperties-member(1))|Updates the range based on a 2D array of cell properties, encapsulating things like font, fill, borders, and alignment.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setcolumnproperties-member(1))|Updates the range based on a single-dimensional array of column properties, encapsulating things like font, fill, borders, and alignment.|
||[setDirty()](/javascript/api/excel/excel.range#excel-excel-range-setdirty-member(1))|Set a range to be recalculated when the next recalculation occurs.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setrowproperties-member(1))|Updates the range based on a single-dimensional array of row properties, encapsulating things like font, fill, borders, and alignment.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[address](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-address-member)|Returns the `RangeAreas` reference in A1-style.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-addresslocal-member)|Returns the `RangeAreas` reference in the user locale.|
||[areaCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areacount-member)|Returns the number of rectangular ranges that comprise this `RangeAreas` object.|
||[areas](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areas-member)|Returns a collection of rectangular ranges that comprise this `RangeAreas` object.|
||[calculate()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-calculate-member(1))|Calculates all cells in the `RangeAreas`.|
||[cellCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-cellcount-member)|Returns the number of cells in the `RangeAreas` object, summing up the cell counts of all of the individual rectangular ranges.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-clear-member(1))|Clears values, format, fill, border, and other properties on each of the areas that comprise this `RangeAreas` object.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-conditionalformats-member)|Returns a collection of conditional formats that intersect with any cells in this `RangeAreas` object.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-convertdatatypetotext-member(1))|Converts all cells in the `RangeAreas` with data types into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-converttolinkeddatatype-member(1))|Converts all cells in the `RangeAreas` into linked data types.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-copyfrom-member(1))|Copies cell data or formatting from the source range or `RangeAreas` to the current `RangeAreas`.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-datavalidation-member)|Returns a data validation object for all ranges in the `RangeAreas`.|
||[format](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-format-member)|Returns a `RangeFormat` object, encapsulating the font, fill, borders, alignment, and other properties for all ranges in the `RangeAreas` object.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirecolumn-member(1))|Returns a `RangeAreas` object that represents the entire columns of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11, H2", it returns a `RangeAreas` that represents columns "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirerow-member(1))|Returns a `RangeAreas` object that represents the entire rows of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11", it returns a `RangeAreas` that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersection-member(1))|Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersectionornullobject-member(1))|Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getoffsetrangeareas-member(1))|Returns a `RangeAreas` object that is shifted by the specific row and column offset.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcells-member(1))|Returns a `RangeAreas` object that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcellsornullobject-member(1))|Returns a `RangeAreas` object that represents all the cells that match the specified type and value.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-gettables-member(1))|Returns a scoped collection of tables that overlap with any range in this `RangeAreas` object.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareas-member(1))|Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareasornullobject-member(1))|Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirecolumn-member)|Specifies if all the ranges on this `RangeAreas` object represent entire columns (e.g., "A:C, Q:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirerow-member)|Specifies if all the ranges on this `RangeAreas` object represent entire rows (e.g., "1:3, 5:7").|
||[setDirty()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-setdirty-member(1))|Sets the `RangeAreas` to be recalculated when the next recalculation occurs.|
||[style](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-style-member)|Represents the style for all ranges in this `RangeAreas` object.|
||[worksheet](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-worksheet-member)|Returns the worksheet for the current `RangeAreas`.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-tintandshade-member)|Specifies a double that lightens or darkens a color for the range border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-tintandshade-member)|Specifies a double that lightens or darkens a color for range borders.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getcount-member(1))|Returns the number of ranges in the `RangeCollection`.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getitemat-member(1))|Returns the range object based on its position in the `RangeCollection`.|
||[items](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-items-member)|Gets the loaded child items in this collection.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-pattern-member)|The pattern of a range.|
||[patternColor](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterncolor-member)|The HTML color code representing the color of the range pattern, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterntintandshade-member)|Specifies a double that lightens or darkens a pattern color for the range fill.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-tintandshade-member)|Specifies a double that lightens or darkens a color for the range fill.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-strikethrough-member)|Specifies the strikethrough status of font.|
||[subscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-subscript-member)|Specifies the subscript status of font.|
||[superscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-superscript-member)|Specifies the superscript status of font.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-tintandshade-member)|Specifies a double that lightens or darkens a color for the range font.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-autoindent-member)|Specifies if text is automatically indented when text alignment is set to equal distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-indentlevel-member)|An integer from 0 to 250 that indicates the indent level.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-readingorder-member)|The reading order for the range.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-shrinktofit-member)|Specifies if text automatically shrinks to fit in the available column width.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-removed-member)|Number of duplicated rows removed by the operation.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-uniqueremaining-member)|Number of remaining unique rows present in the resulting range.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-completematch-member)|Specifies if the match needs to be complete or partial.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-matchcase-member)|Specifies if the match is case-sensitive.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-address-member)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-addresslocal-member)|Represents the `addressLocal` property.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-rowindex-member)|Represents the `rowIndex` property.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-completematch-member)|Specifies if the match needs to be complete or partial.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-matchcase-member)|Specifies if the match is case-sensitive.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-searchdirection-member)|Specifies the search direction.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-format-member)|Represents the `format` property.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-hyperlink-member)|Represents the `hyperlink` property.|
||[style](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-style-member)|Represents the `style` property.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnhidden-member)|Represents the `columnHidden` property.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnwidth-member)||
||[format: Excel.CellPropertiesFormat & {            columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-format-member)|Represents the `format` property.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel.CellPropertiesFormat & {            rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-format-member)|Represents the `format` property.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowheight-member)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowhidden-member)|Represents the `rowHidden` property.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#excel-excel-shape-alttextdescription-member)|Specifies the alternative description text for a `Shape` object.|
||[altTextTitle](/javascript/api/excel/excel.shape#excel-excel-shape-alttexttitle-member)|Specifies the alternative title text for a `Shape` object.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#excel-excel-shape-connectionsitecount-member)|Returns the number of connection sites on this shape.|
||[delete()](/javascript/api/excel/excel.shape#excel-excel-shape-delete-member(1))|Removes the shape from the worksheet.|
||[fill](/javascript/api/excel/excel.shape#excel-excel-shape-fill-member)|Returns the fill formatting of this shape.|
||[geometricShape](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshape-member)|Returns the geometric shape associated with the shape.|
||[geometricShapeType](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshapetype-member)|Specifies the geometric shape type of this geometric shape.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1))|Converts the shape to an image and returns the image as a base64-encoded string.|
||[group](/javascript/api/excel/excel.shape#excel-excel-shape-group-member)|Returns the shape group associated with the shape.|
||[height](/javascript/api/excel/excel.shape#excel-excel-shape-height-member)|Specifies the height, in points, of the shape.|
||[id](/javascript/api/excel/excel.shape#excel-excel-shape-id-member)|Specifies the shape identifier.|
||[image](/javascript/api/excel/excel.shape#excel-excel-shape-image-member)|Returns the image associated with the shape.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementleft-member(1))|Moves the shape horizontally by the specified number of points.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementrotation-member(1))|Rotates the shape clockwise around the z-axis by the specified number of degrees.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementtop-member(1))|Moves the shape vertically by the specified number of points.|
||[left](/javascript/api/excel/excel.shape#excel-excel-shape-left-member)|The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[level](/javascript/api/excel/excel.shape#excel-excel-shape-level-member)|Specifies the level of the specified shape.|
||[line](/javascript/api/excel/excel.shape#excel-excel-shape-line-member)|Returns the line associated with the shape.|
||[lineFormat](/javascript/api/excel/excel.shape#excel-excel-shape-lineformat-member)|Returns the line formatting of this shape.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#excel-excel-shape-lockaspectratio-member)|Specifies if the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.shape#excel-excel-shape-name-member)|Specifies the name of the shape.|
||[onActivated](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)|Occurs when the shape is activated.|
||[onDeactivated](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)|Occurs when the shape is deactivated.|
||[parentGroup](/javascript/api/excel/excel.shape#excel-excel-shape-parentgroup-member)|Specifies the parent group of this shape.|
||[rotation](/javascript/api/excel/excel.shape#excel-excel-shape-rotation-member)|Specifies the rotation, in degrees, of the shape.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scaleheight-member(1))|Scales the height of the shape by a specified factor.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scalewidth-member(1))|Scales the width of the shape by a specified factor.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#excel-excel-shape-setzorder-member(1))|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[textFrame](/javascript/api/excel/excel.shape#excel-excel-shape-textframe-member)|Returns the text frame object of this shape.|
||[top](/javascript/api/excel/excel.shape#excel-excel-shape-top-member)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[type](/javascript/api/excel/excel.shape#excel-excel-shape-type-member)|Returns the type of this shape.|
||[visible](/javascript/api/excel/excel.shape#excel-excel-shape-visible-member)|Specifies if the shape is visible.|
||[width](/javascript/api/excel/excel.shape#excel-excel-shape-width-member)|Specifies the width, in points, of the shape.|
||[zOrderPosition](/javascript/api/excel/excel.shape#excel-excel-shape-zorderposition-member)|Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-shapeid-member)|Gets the ID of the activated shape.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the shape is activated.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1))|Adds a geometric shape to the worksheet.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgroup-member(1))|Groups a subset of shapes in this collection's worksheet.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1))|Creates an image from a base64-encoded string and adds it to the worksheet.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1))|Adds a line to worksheet.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1))|Adds a text box to the worksheet with the provided text as the content.|
||[getCount()](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getcount-member(1))|Returns the number of shapes in the worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitem-member(1))|Gets a shape using its name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemat-member(1))|Gets a shape using its position in the collection.|
||[items](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-items-member)|Gets the loaded child items in this collection.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-shapeid-member)|Gets the ID of the shape deactivated shape.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the shape is deactivated.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-clear-member(1))|Clears the fill formatting of this shape.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-foregroundcolor-member)|Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange")|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-setsolidcolor-member(1))|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-type-member)|Returns the fill type of the shape.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-bold-member)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-color-member)|HTML color code representation of the text color (e.g., "#FF0000" represents red).|
||[italic](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-italic-member)|Represents the italic status of font.|
||[name](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-name-member)|Represents font name (e.g., "Calibri").|
||[size](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-size-member)|Represents font size in points (e.g., 11).|
||[underline](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-underline-member)|Type of underline applied to the font.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-id-member)|Specifies the shape identifier.|
||[shape](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shape-member)|Returns the `Shape` object associated with the group.|
||[shapes](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shapes-member)|Returns the collection of `Shape` objects.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-color-member)|Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-dashstyle-member)|Represents the line style of the shape.|
||[style](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-style-member)|Represents the line style of the shape.|
||[transparency](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-transparency-member)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear).|
||[visible](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-visible-member)|Specifies if the line formatting of a shape element is visible.|
||[weight](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-weight-member)|Represents the weight of the line, in points.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#excel-excel-sortfield-subfield-member)|Specifies the subfield that is the target property name of a rich value to sort on.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getcount-member(1))|Gets the number of styles in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemat-member(1))|Gets a style based on its position in the collection.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#excel-excel-table-autofilter-member)|Represents the `AutoFilter` object of the table.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-source-member)|Gets the source of the event.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-tableid-member)|Gets the ID of the table that is added.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the table is added.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-details-member)|Gets the information about the change detail.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member)|Occurs when a new table is added in a workbook.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member)|Occurs when the specified table is deleted in a workbook.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-source-member)|Gets the source of the event.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tableid-member)|Gets the ID of the table that is deleted.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tablename-member)|Gets the name of the table that is deleted.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the table is deleted.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getcount-member(1))|Gets the number of tables in the collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getfirst-member(1))|Gets the first table in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitem-member(1))|Gets a table by name or ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#excel-excel-textframe-autosizesetting-member)|The automatic sizing settings for the text frame.|
||[bottomMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-bottommargin-member)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/javascript/api/excel/excel.textframe#excel-excel-textframe-deletetext-member(1))|Deletes all the text in the text frame.|
||[hasText](/javascript/api/excel/excel.textframe#excel-excel-textframe-hastext-member)|Specifies if the text frame contains text.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontalalignment-member)|Represents the horizontal alignment of the text frame.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontaloverflow-member)|Represents the horizontal overflow behavior of the text frame.|
||[leftMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-leftmargin-member)|Represents the left margin, in points, of the text frame.|
||[orientation](/javascript/api/excel/excel.textframe#excel-excel-textframe-orientation-member)|Represents the angle to which the text is oriented for the text frame.|
||[readingOrder](/javascript/api/excel/excel.textframe#excel-excel-textframe-readingorder-member)|Represents the reading order of the text frame, either left-to-right or right-to-left.|
||[rightMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-rightmargin-member)|Represents the right margin, in points, of the text frame.|
||[textRange](/javascript/api/excel/excel.textframe#excel-excel-textframe-textrange-member)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text.|
||[topMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-topmargin-member)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticalalignment-member)|Represents the vertical alignment of the text frame.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticaloverflow-member)|Represents the vertical overflow behavior of the text frame.|
|[TextRange](/javascript/api/excel/excel.textrange)|[font](/javascript/api/excel/excel.textrange#excel-excel-textrange-font-member)|Returns a `ShapeFont` object that represents the font attributes for the text range.|
||[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#excel-excel-textrange-getsubstring-member(1))|Returns a TextRange object for the substring in the given range.|
||[text](/javascript/api/excel/excel.textrange#excel-excel-textrange-text-member)|Represents the plain text content of the text range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[autoSave](/javascript/api/excel/excel.workbook#excel-excel-workbook-autosave-member)|Specifies if the workbook is in AutoSave mode.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#excel-excel-workbook-calculationengineversion-member)|Returns a number about the version of Excel Calculation Engine.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbook#excel-excel-workbook-chartdatapointtrack-member)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechart-member(1))|Gets the currently active chart in the workbook.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechartornullobject-member(1))|Gets the currently active chart in the workbook.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getisactivecollabsession-member(1))|Returns `true` if the workbook is being edited by multiple users (through co-authoring).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedranges-member(1))|Gets the currently selected one or more ranges from the workbook.|
||[isDirty](/javascript/api/excel/excel.workbook#excel-excel-workbook-isdirty-member)|Specifies if changes have been made since the workbook was last saved.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member)|Occurs when the AutoSave setting is changed on the workbook.|
||[previouslySaved](/javascript/api/excel/excel.workbook#excel-excel-workbook-previouslysaved-member)|Specifies if the workbook has ever been saved locally or online.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#excel-excel-workbook-useprecisionasdisplayed-member)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#excel-excel-workbookautosavesettingchangedeventargs-type-member)|Gets the type of the event.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[autoFilter](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-autofilter-member)|Represents the `AutoFilter` object of the worksheet.|
||[enableCalculation](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-enablecalculation-member)|Determines if Excel should recalculate the worksheet when necessary.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findall-member(1))|Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findallornullobject-member(1))|Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1))|Gets the `RangeAreas` object, representing one or more blocks of rectangular ranges, specified by the address or name.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-horizontalpagebreaks-member)|Gets the horizontal page break collection for the worksheet.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)|Occurs when format changed on a specific worksheet.|
||[pageLayout](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pagelayout-member)|Gets the `PageLayout` object of the worksheet.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-replaceall-member(1))|Finds and replaces the given string based on the criteria specified within the current worksheet.|
||[shapes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-shapes-member)|Returns the collection of all the Shape objects on the worksheet.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-verticalpagebreaks-member)|Gets the vertical page break collection for the worksheet.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-details-member)|Represents the information about the change detail.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member)|Occurs when any worksheet in the workbook is changed.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member)|Occurs when any worksheet in the workbook has a format changed.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member)|Occurs when the selection changes on any worksheet.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-address-member)|Gets the range address that represents the changed area of a specific worksheet.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrange-member(1))|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrangeornullobject-member(1))|Gets the range that represents the changed area of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-completematch-member)|Specifies if the match needs to be complete or partial.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-matchcase-member)|Specifies if the match is case-sensitive.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
