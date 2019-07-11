---
title: Excel JavaScript API requirement set 1.9
description: 'Details about the ExcelApi 1.9 requirement set'
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
---

# What’s new in Excel JavaScript API 1.9

More than 500 new Excel APIs were introduced with the 1.9 requirement set. The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Shapes](../../../excel/excel-add-ins-shapes.md) | Insert, position, and format images, geometric shapes and text boxes. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Auto Filter](../../../excel/excel-add-ins-worksheets.md#filter-data) | Add filters to ranges. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](../../../excel/excel-add-ins-multiple-ranges.md) | Support for discontinuous ranges. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Special Cells](../../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Get cells containing dates, comments, or formulas within a range. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Find](../../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | Find values or formulas within a range or worksheet. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copy and Paste](../../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | Copy values, formats, and formulas from one range to another. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](../../../excel/performance.md#suspend-calculation-temporarily) | Greater control over the Excel calculation engine. | [Application](/javascript/api/excel/excel.application) |
| New Charts | Explore our new supported chart types: maps, box and whisker, waterfall, sunburst, pareto. and funnel. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | New capabilities with range formats. | [Range](/javascript/api/excel/excel.rangeformat) |

## API list

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Returns the Excel calculation engine version used for the last full recalculation. Read-only.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Returns the calculation state of the application. See Excel.CalculationState for details. Read-only.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Returns the Iterative Calculation settings.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Suspends sceen updating until the next "context.sync()" is called.|
|[ApplicationData](/javascript/api/excel/excel.applicationdata)|[calculationEngineVersion](/javascript/api/excel/excel.applicationdata#calculationengineversion)|Returns the Excel calculation engine version used for the last full recalculation. Read-only.|
||[calculationState](/javascript/api/excel/excel.applicationdata#calculationstate)|Returns the calculation state of the application. See Excel.CalculationState for details. Read-only.|
||[iterativeCalculation](/javascript/api/excel/excel.applicationdata#iterativecalculation)|Returns the Iterative Calculation settings.|
|[ApplicationLoadOptions](/javascript/api/excel/excel.applicationloadoptions)|[calculationEngineVersion](/javascript/api/excel/excel.applicationloadoptions#calculationengineversion)|Returns the Excel calculation engine version used for the last full recalculation. Read-only.|
||[calculationState](/javascript/api/excel/excel.applicationloadoptions#calculationstate)|Returns the calculation state of the application. See Excel.CalculationState for details. Read-only.|
||[iterativeCalculation](/javascript/api/excel/excel.applicationloadoptions#iterativecalculation)|Returns the Iterative Calculation settings.|
|[ApplicationUpdateData](/javascript/api/excel/excel.applicationupdatedata)|[iterativeCalculation](/javascript/api/excel/excel.applicationupdatedata#iterativecalculation)|Returns the Iterative Calculation settings.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Applies the AutoFilter to a range. This filters the column if column index and filter criteria are specified.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Clears the filter criteria of the AutoFilter.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Returns the Range object that represents the range to which the AutoFilter applies.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Returns the Range object that represents the range to which the AutoFilter applies.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|An array that holds all the filter criteria in the autofiltered range. Read-Only.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Indicates if the AutoFilter is enabled or not. Read-Only.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Indicates if the AutoFilter has filter criteria. Read-Only.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Applies the specified Autofilter object currently on the range.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Removes the AutoFilter for the range.|
|[AutoFilterData](/javascript/api/excel/excel.autofilterdata)|[criteria](/javascript/api/excel/excel.autofilterdata#criteria)|An array that holds all the filter criteria in the autofiltered range. Read-Only.|
||[enabled](/javascript/api/excel/excel.autofilterdata#enabled)|Indicates if the AutoFilter is enabled or not. Read-Only.|
||[isDataFiltered](/javascript/api/excel/excel.autofilterdata#isdatafiltered)|Indicates if the AutoFilter has filter criteria. Read-Only.|
|[AutoFilterLoadOptions](/javascript/api/excel/excel.autofilterloadoptions)|[$all](/javascript/api/excel/excel.autofilterloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.autofilterloadoptions#criteria)|An array that holds all the filter criteria in the autofiltered range. Read-Only.|
||[enabled](/javascript/api/excel/excel.autofilterloadoptions#enabled)|Indicates if the AutoFilter is enabled or not. Read-Only.|
||[isDataFiltered](/javascript/api/excel/excel.autofilterloadoptions#isdatafiltered)|Indicates if the AutoFilter has filter criteria. Read-Only.|
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
|[CellPropertiesBorderLoadOptions](/javascript/api/excel/excel.cellpropertiesborderloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesborderloadoptions#color)|Specifies whether to load on the `color` property.|
||[style](/javascript/api/excel/excel.cellpropertiesborderloadoptions#style)|Specifies whether to load on the `style` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesborderloadoptions#tintandshade)|Specifies whether to load on the `tintAndShade` property.|
||[weight](/javascript/api/excel/excel.cellpropertiesborderloadoptions#weight)|Specifies whether to load on the `weight` property.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Represents the `format.fill.color` property.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Represents the `format.fill.pattern` property.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Represents the `format.fill.patternColor` property.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Represents the `format.fill.patternTintAndShade` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Represents the `format.fill.tintAndShade` property.|
|[CellPropertiesFillLoadOptions](/javascript/api/excel/excel.cellpropertiesfillloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesfillloadoptions#color)|Specifies whether to load on the `color` property.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfillloadoptions#pattern)|Specifies whether to load on the `pattern` property.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterncolor)|Specifies whether to load on the `patternColor` property.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterntintandshade)|Specifies whether to load on the `patternTintAndShade` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#tintandshade)|Specifies whether to load on the `tintAndShade` property.|
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
|[CellPropertiesFontLoadOptions](/javascript/api/excel/excel.cellpropertiesfontloadoptions)|[bold](/javascript/api/excel/excel.cellpropertiesfontloadoptions#bold)|Specifies whether to load on the `bold` property.|
||[color](/javascript/api/excel/excel.cellpropertiesfontloadoptions#color)|Specifies whether to load on the `color` property.|
||[italic](/javascript/api/excel/excel.cellpropertiesfontloadoptions#italic)|Specifies whether to load on the `italic` property.|
||[name](/javascript/api/excel/excel.cellpropertiesfontloadoptions#name)|Specifies whether to load on the `name` property.|
||[size](/javascript/api/excel/excel.cellpropertiesfontloadoptions#size)|Specifies whether to load on the `size` property.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfontloadoptions#strikethrough)|Specifies whether to load on the `strikethrough` property.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#subscript)|Specifies whether to load on the `subscript` property.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#superscript)|Specifies whether to load on the `superscript` property.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfontloadoptions#tintandshade)|Specifies whether to load on the `tintAndShade` property.|
||[underline](/javascript/api/excel/excel.cellpropertiesfontloadoptions#underline)|Specifies whether to load on the `underline` property.|
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
|[CellPropertiesFormatLoadOptions](/javascript/api/excel/excel.cellpropertiesformatloadoptions)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformatloadoptions#autoindent)|Specifies whether to load on the `autoIndent` property.|
||[borders](/javascript/api/excel/excel.cellpropertiesformatloadoptions#borders)|Specifies whether to load on the `borders` property.|
||[fill](/javascript/api/excel/excel.cellpropertiesformatloadoptions#fill)|Specifies whether to load on the `fill` property.|
||[font](/javascript/api/excel/excel.cellpropertiesformatloadoptions#font)|Specifies whether to load on the `font` property.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#horizontalalignment)|Specifies whether to load on the `horizontalAlignment` property.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformatloadoptions#indentlevel)|Specifies whether to load on the `indentLevel` property.|
||[protection](/javascript/api/excel/excel.cellpropertiesformatloadoptions#protection)|Specifies whether to load on the `protection` property.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformatloadoptions#readingorder)|Specifies whether to load on the `readingOrder` property.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformatloadoptions#shrinktofit)|Specifies whether to load on the `shrinkToFit` property.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformatloadoptions#textorientation)|Specifies whether to load on the `textOrientation` property.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardheight)|Specifies whether to load on the `useStandardHeight` property.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardwidth)|Specifies whether to load on the `useStandardWidth` property.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#verticalalignment)|Specifies whether to load on the `verticalAlignment` property.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformatloadoptions#wraptext)|Specifies whether to load on the `wrapText` property.|
|[CellPropertiesLoadOptions](/javascript/api/excel/excel.cellpropertiesloadoptions)|[address](/javascript/api/excel/excel.cellpropertiesloadoptions#address)|Specifies whether to load on the `address` property.|
||[addressLocal](/javascript/api/excel/excel.cellpropertiesloadoptions#addresslocal)|Specifies whether to load on the `addressLocal` property.|
||[format](/javascript/api/excel/excel.cellpropertiesloadoptions#format)|Specifies whether to load on the `format` property.|
||[hidden](/javascript/api/excel/excel.cellpropertiesloadoptions#hidden)|Specifies whether to load on the `hidden` property.|
||[hyperlink](/javascript/api/excel/excel.cellpropertiesloadoptions#hyperlink)|Specifies whether to load on the `hyperlink` property.|
||[style](/javascript/api/excel/excel.cellpropertiesloadoptions#style)|Specifies whether to load on the `style` property.|
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
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[colorScheme](/javascript/api/excel/excel.chartareaformatdata#colorscheme)|Returns or sets color scheme of the chart. Read/Write.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatdata#roundedcorners)|Specifies whether or not chart area of the chart has rounded corners. Read/Write.|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[colorScheme](/javascript/api/excel/excel.chartareaformatloadoptions#colorscheme)|Returns or sets color scheme of the chart. Read/Write.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatloadoptions#roundedcorners)|Specifies whether or not chart area of the chart has rounded corners. Read/Write.|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[colorScheme](/javascript/api/excel/excel.chartareaformatupdatedata#colorscheme)|Returns or sets color scheme of the chart. Read/Write.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatupdatedata#roundedcorners)|Specifies whether or not chart area of the chart has rounded corners. Read/Write.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells.|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisdata#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisloadoptions#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisupdatedata#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Specifies whether or not the bin overflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Specifies whether or not the bin underflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Returns or sets the bin count of a histogram chart or pareto chart. Read/Write.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Returns or sets the bin overflow value of a histogram chart or pareto chart. Read/Write.|
||[set(properties: Excel.ChartBinOptions)](/javascript/api/excel/excel.chartbinoptions#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartBinOptionsUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartbinoptions#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Returns or sets the bin's type for a histogram chart or pareto chart. Read/Write.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Returns or sets the bin underflow value of a histogram chart or pareto chart. Read/Write.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Returns or sets the bin width value of a histogram chart or pareto chart. Read/Write.|
|[ChartBinOptionsData](/javascript/api/excel/excel.chartbinoptionsdata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsdata#allowoverflow)|Specifies whether or not the bin overflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsdata#allowunderflow)|Specifies whether or not the bin underflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[count](/javascript/api/excel/excel.chartbinoptionsdata#count)|Returns or sets the bin count of a histogram chart or pareto chart. Read/Write.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsdata#overflowvalue)|Returns or sets the bin overflow value of a histogram chart or pareto chart. Read/Write.|
||[type](/javascript/api/excel/excel.chartbinoptionsdata#type)|Returns or sets the bin's type for a histogram chart or pareto chart. Read/Write.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsdata#underflowvalue)|Returns or sets the bin underflow value of a histogram chart or pareto chart. Read/Write.|
||[width](/javascript/api/excel/excel.chartbinoptionsdata#width)|Returns or sets the bin width value of a histogram chart or pareto chart. Read/Write.|
|[ChartBinOptionsLoadOptions](/javascript/api/excel/excel.chartbinoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartbinoptionsloadoptions#$all)||
||[allowOverflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowoverflow)|Specifies whether or not the bin overflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowunderflow)|Specifies whether or not the bin underflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[count](/javascript/api/excel/excel.chartbinoptionsloadoptions#count)|Returns or sets the bin count of a histogram chart or pareto chart. Read/Write.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#overflowvalue)|Returns or sets the bin overflow value of a histogram chart or pareto chart. Read/Write.|
||[type](/javascript/api/excel/excel.chartbinoptionsloadoptions#type)|Returns or sets the bin's type for a histogram chart or pareto chart. Read/Write.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#underflowvalue)|Returns or sets the bin underflow value of a histogram chart or pareto chart. Read/Write.|
||[width](/javascript/api/excel/excel.chartbinoptionsloadoptions#width)|Returns or sets the bin width value of a histogram chart or pareto chart. Read/Write.|
|[ChartBinOptionsUpdateData](/javascript/api/excel/excel.chartbinoptionsupdatedata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowoverflow)|Specifies whether or not the bin overflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowunderflow)|Specifies whether or not the bin underflow is enabled in a histogram chart or pareto chart. Read/Write.|
||[count](/javascript/api/excel/excel.chartbinoptionsupdatedata#count)|Returns or sets the bin count of a histogram chart or pareto chart. Read/Write.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#overflowvalue)|Returns or sets the bin overflow value of a histogram chart or pareto chart. Read/Write.|
||[type](/javascript/api/excel/excel.chartbinoptionsupdatedata#type)|Returns or sets the bin's type for a histogram chart or pareto chart. Read/Write.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#underflowvalue)|Returns or sets the bin underflow value of a histogram chart or pareto chart. Read/Write.|
||[width](/javascript/api/excel/excel.chartbinoptionsupdatedata#width)|Returns or sets the bin width value of a histogram chart or pareto chart. Read/Write.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Returns or sets the quartile calculation type of a box and whisker chart. Read/Write.|
||[set(properties: Excel.ChartBoxwhiskerOptions)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartBoxwhiskerOptionsUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Specifies whether or not the inner points are shown in a box and whisker chart. Read/Write.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Specifies whether or not the mean line is shown in a box and whisker chart. Read/Write.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Specifies whether or not the mean marker is shown in a box and whisker chart. Read/Write.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Specifies whether or not outlier points are shown in a box and whisker chart. Read/Write.|
|[ChartBoxwhiskerOptionsData](/javascript/api/excel/excel.chartboxwhiskeroptionsdata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#quartilecalculation)|Returns or sets the quartile calculation type of a box and whisker chart. Read/Write.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showinnerpoints)|Specifies whether or not the inner points are shown in a box and whisker chart. Read/Write.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanline)|Specifies whether or not the mean line is shown in a box and whisker chart. Read/Write.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanmarker)|Specifies whether or not the mean marker is shown in a box and whisker chart. Read/Write.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showoutlierpoints)|Specifies whether or not outlier points are shown in a box and whisker chart. Read/Write.|
|[ChartBoxwhiskerOptionsLoadOptions](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions)|[$all](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#$all)||
||[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#quartilecalculation)|Returns or sets the quartile calculation type of a box and whisker chart. Read/Write.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showinnerpoints)|Specifies whether or not the inner points are shown in a box and whisker chart. Read/Write.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanline)|Specifies whether or not the mean line is shown in a box and whisker chart. Read/Write.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanmarker)|Specifies whether or not the mean marker is shown in a box and whisker chart. Read/Write.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showoutlierpoints)|Specifies whether or not outlier points are shown in a box and whisker chart. Read/Write.|
|[ChartBoxwhiskerOptionsUpdateData](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#quartilecalculation)|Returns or sets the quartile calculation type of a box and whisker chart. Read/Write.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showinnerpoints)|Specifies whether or not the inner points are shown in a box and whisker chart. Read/Write.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanline)|Specifies whether or not the mean line is shown in a box and whisker chart. Read/Write.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanmarker)|Specifies whether or not the mean marker is shown in a box and whisker chart. Read/Write.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showoutlierpoints)|Specifies whether or not outlier points are shown in a box and whisker chart. Read/Write.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartcollectionloadoptions#pivotoptions)|For EACH ITEM in the collection: Encapsulates the options for a pivot chart.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[pivotOptions](/javascript/api/excel/excel.chartdata#pivotoptions)|Encapsulates the options for a pivot chart. Read-only.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabeldata#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsdata#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#linknumberformat)|Represents whether or not the number format is linked to the cells. If true, the number format will change in the labels when it changes in the cells|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Specifies whether or not the error bars have an end style cap.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Specifies which parts of the error bars to include.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Specifies the formatting type of the error bars.|
||[set(properties: Excel.ChartErrorBars)](/javascript/api/excel/excel.charterrorbars#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartErrorBarsUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.charterrorbars#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|The type of range marked by the error bars.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Specifies whether or not the error bars are displayed.|
|[ChartErrorBarsData](/javascript/api/excel/excel.charterrorbarsdata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsdata#endstylecap)|Specifies whether or not the error bars have an end style cap.|
||[format](/javascript/api/excel/excel.charterrorbarsdata#format)|Specifies the formatting type of the error bars.|
||[include](/javascript/api/excel/excel.charterrorbarsdata#include)|Specifies which parts of the error bars to include.|
||[type](/javascript/api/excel/excel.charterrorbarsdata#type)|The type of range marked by the error bars.|
||[visible](/javascript/api/excel/excel.charterrorbarsdata#visible)|Specifies whether or not the error bars are displayed.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Represents the chart line formatting.|
||[set(properties: Excel.ChartErrorBarsFormat)](/javascript/api/excel/excel.charterrorbarsformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartErrorBarsFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.charterrorbarsformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartErrorBarsFormatData](/javascript/api/excel/excel.charterrorbarsformatdata)|[line](/javascript/api/excel/excel.charterrorbarsformatdata#line)|Represents the chart line formatting.|
|[ChartErrorBarsFormatLoadOptions](/javascript/api/excel/excel.charterrorbarsformatloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charterrorbarsformatloadoptions#line)|Represents the chart line formatting.|
|[ChartErrorBarsFormatUpdateData](/javascript/api/excel/excel.charterrorbarsformatupdatedata)|[line](/javascript/api/excel/excel.charterrorbarsformatupdatedata#line)|Represents the chart line formatting.|
|[ChartErrorBarsLoadOptions](/javascript/api/excel/excel.charterrorbarsloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsloadoptions#$all)||
||[endStyleCap](/javascript/api/excel/excel.charterrorbarsloadoptions#endstylecap)|Specifies whether or not the error bars have an end style cap.|
||[format](/javascript/api/excel/excel.charterrorbarsloadoptions#format)|Specifies the formatting type of the error bars.|
||[include](/javascript/api/excel/excel.charterrorbarsloadoptions#include)|Specifies which parts of the error bars to include.|
||[type](/javascript/api/excel/excel.charterrorbarsloadoptions#type)|The type of range marked by the error bars.|
||[visible](/javascript/api/excel/excel.charterrorbarsloadoptions#visible)|Specifies whether or not the error bars are displayed.|
|[ChartErrorBarsUpdateData](/javascript/api/excel/excel.charterrorbarsupdatedata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsupdatedata#endstylecap)|Specifies whether or not the error bars have an end style cap.|
||[format](/javascript/api/excel/excel.charterrorbarsupdatedata#format)|Specifies the formatting type of the error bars.|
||[include](/javascript/api/excel/excel.charterrorbarsupdatedata#include)|Specifies which parts of the error bars to include.|
||[type](/javascript/api/excel/excel.charterrorbarsupdatedata#type)|The type of range marked by the error bars.|
||[visible](/javascript/api/excel/excel.charterrorbarsupdatedata#visible)|Specifies whether or not the error bars are displayed.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartloadoptions#pivotoptions)|Encapsulates the options for a pivot chart.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Returns or sets the series map labels strategy of a region map chart. Read/Write.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Returns or sets the series mapping level of a region map chart. Read/Write.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Returns or sets the series projection type of a region map chart. Read/Write.|
||[set(properties: Excel.ChartMapOptions)](/javascript/api/excel/excel.chartmapoptions#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartMapOptionsUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartmapoptions#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ChartMapOptionsData](/javascript/api/excel/excel.chartmapoptionsdata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsdata#labelstrategy)|Returns or sets the series map labels strategy of a region map chart. Read/Write.|
||[level](/javascript/api/excel/excel.chartmapoptionsdata#level)|Returns or sets the series mapping level of a region map chart. Read/Write.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsdata#projectiontype)|Returns or sets the series projection type of a region map chart. Read/Write.|
|[ChartMapOptionsLoadOptions](/javascript/api/excel/excel.chartmapoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartmapoptionsloadoptions#$all)||
||[labelStrategy](/javascript/api/excel/excel.chartmapoptionsloadoptions#labelstrategy)|Returns or sets the series map labels strategy of a region map chart. Read/Write.|
||[level](/javascript/api/excel/excel.chartmapoptionsloadoptions#level)|Returns or sets the series mapping level of a region map chart. Read/Write.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsloadoptions#projectiontype)|Returns or sets the series projection type of a region map chart. Read/Write.|
|[ChartMapOptionsUpdateData](/javascript/api/excel/excel.chartmapoptionsupdatedata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsupdatedata#labelstrategy)|Returns or sets the series map labels strategy of a region map chart. Read/Write.|
||[level](/javascript/api/excel/excel.chartmapoptionsupdatedata#level)|Returns or sets the series mapping level of a region map chart. Read/Write.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsupdatedata#projectiontype)|Returns or sets the series projection type of a region map chart. Read/Write.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[set(properties: Excel.ChartPivotOptions)](/javascript/api/excel/excel.chartpivotoptions#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ChartPivotOptionsUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.chartpivotoptions#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Specifies whether or not to display the axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Specifies whether or not to display the legend field buttons on a PivotChart|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Specifies whether or not to display the report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Specifies whether or not to display the show value field buttons on a PivotChart|
|[ChartPivotOptionsData](/javascript/api/excel/excel.chartpivotoptionsdata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showaxisfieldbuttons)|Specifies whether or not to display the axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showlegendfieldbuttons)|Specifies whether or not to display the legend field buttons on a PivotChart|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showreportfilterfieldbuttons)|Specifies whether or not to display the report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showvaluefieldbuttons)|Specifies whether or not to display the show value field buttons on a PivotChart|
|[ChartPivotOptionsLoadOptions](/javascript/api/excel/excel.chartpivotoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartpivotoptionsloadoptions#$all)||
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showaxisfieldbuttons)|Specifies whether or not to display the axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showlegendfieldbuttons)|Specifies whether or not to display the legend field buttons on a PivotChart|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showreportfilterfieldbuttons)|Specifies whether or not to display the report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showvaluefieldbuttons)|Specifies whether or not to display the show value field buttons on a PivotChart|
|[ChartPivotOptionsUpdateData](/javascript/api/excel/excel.chartpivotoptionsupdatedata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showaxisfieldbuttons)|Specifies whether or not to display the axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the "Show Axis Field Buttons" command on the "Field Buttons" drop-down list of the "Analyze" tab, which is available when a PivotChart is selected.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showlegendfieldbuttons)|Specifies whether or not to display the legend field buttons on a PivotChart|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showreportfilterfieldbuttons)|Specifies whether or not to display the report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showvaluefieldbuttons)|Specifies whether or not to display the show value field buttons on a PivotChart|
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
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#binoptions)|For EACH ITEM in the collection: Encapsulates the bin options for histogram charts and pareto charts.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#boxwhiskeroptions)|For EACH ITEM in the collection: Encapsulates the options for the box and whisker charts.|
||[bubbleScale](/javascript/api/excel/excel.chartseriescollectionloadoptions#bubblescale)|For EACH ITEM in the collection: This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts. Read/Write.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumcolor)|For EACH ITEM in the collection: Returns or sets the color for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumtype)|For EACH ITEM in the collection: Returns or sets the type for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumvalue)|For EACH ITEM in the collection: Returns or sets the maximum value of a region map chart series. Read/Write.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointcolor)|For EACH ITEM in the collection: Returns or sets the color for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointtype)|For EACH ITEM in the collection: Returns or sets the type for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointvalue)|For EACH ITEM in the collection: Returns or sets the midpoint value of a region map chart series. Read/Write.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumcolor)|For EACH ITEM in the collection: Returns or sets the color for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumtype)|For EACH ITEM in the collection: Returns or sets the type for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumvalue)|For EACH ITEM in the collection: Returns or sets the minimum value of a region map chart series. Read/Write.|
||[gradientStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientstyle)|For EACH ITEM in the collection: Returns or sets series gradient style of a region map chart. Read/Write.|
||[invertColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertcolor)|For EACH ITEM in the collection: Returns or sets the fill color for negative data points in a series. Read/Write.|
||[mapOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#mapoptions)|For EACH ITEM in the collection: Encapsulates the options for a region map chart.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriescollectionloadoptions#parentlabelstrategy)|For EACH ITEM in the collection: Returns or sets the series parent label strategy area for a treemap chart. Read/Write.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showconnectorlines)|For EACH ITEM in the collection: Specifies whether or not connector lines are shown in waterfall charts. Read/Write.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showleaderlines)|For EACH ITEM in the collection: Specifies whether or not leader lines are displayed for each data label in the series. Read/Write.|
||[splitValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#splitvalue)|For EACH ITEM in the collection: Returns or sets the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/Write.|
||[xErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#xerrorbars)|For EACH ITEM in the collection: Represents the error bar object of a chart series.|
||[yErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#yerrorbars)|For EACH ITEM in the collection: Represents the error bar object of a chart series.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[binOptions](/javascript/api/excel/excel.chartseriesdata#binoptions)|Encapsulates the bin options for histogram charts and pareto charts. Read-only.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesdata#boxwhiskeroptions)|Encapsulates the options for the box and whisker charts. Read-only.|
||[bubbleScale](/javascript/api/excel/excel.chartseriesdata#bubblescale)|This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts. Read/Write.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesdata#gradientmaximumcolor)|Returns or sets the color for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesdata#gradientmaximumtype)|Returns or sets the type for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesdata#gradientmaximumvalue)|Returns or sets the maximum value of a region map chart series. Read/Write.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesdata#gradientmidpointcolor)|Returns or sets the color for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesdata#gradientmidpointtype)|Returns or sets the type for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesdata#gradientmidpointvalue)|Returns or sets the midpoint value of a region map chart series. Read/Write.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesdata#gradientminimumcolor)|Returns or sets the color for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesdata#gradientminimumtype)|Returns or sets the type for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesdata#gradientminimumvalue)|Returns or sets the minimum value of a region map chart series. Read/Write.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesdata#gradientstyle)|Returns or sets series gradient style of a region map chart. Read/Write.|
||[invertColor](/javascript/api/excel/excel.chartseriesdata#invertcolor)|Returns or sets the fill color for negative data points in a series. Read/Write.|
||[mapOptions](/javascript/api/excel/excel.chartseriesdata#mapoptions)|Encapsulates the options for a region map chart. Read-only.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesdata#parentlabelstrategy)|Returns or sets the series parent label strategy area for a treemap chart. Read/Write.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesdata#showconnectorlines)|Specifies whether or not connector lines are shown in waterfall charts. Read/Write.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesdata#showleaderlines)|Specifies whether or not leader lines are displayed for each data label in the series. Read/Write.|
||[splitValue](/javascript/api/excel/excel.chartseriesdata#splitvalue)|Returns or sets the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/Write.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesdata#xerrorbars)|Represents the error bar object of a chart series.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesdata#yerrorbars)|Represents the error bar object of a chart series.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriesloadoptions#binoptions)|Encapsulates the bin options for histogram charts and pareto charts.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesloadoptions#boxwhiskeroptions)|Encapsulates the options for the box and whisker charts.|
||[bubbleScale](/javascript/api/excel/excel.chartseriesloadoptions#bubblescale)|This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts. Read/Write.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumcolor)|Returns or sets the color for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumtype)|Returns or sets the type for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumvalue)|Returns or sets the maximum value of a region map chart series. Read/Write.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointcolor)|Returns or sets the color for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointtype)|Returns or sets the type for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointvalue)|Returns or sets the midpoint value of a region map chart series. Read/Write.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumcolor)|Returns or sets the color for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumtype)|Returns or sets the type for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumvalue)|Returns or sets the minimum value of a region map chart series. Read/Write.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesloadoptions#gradientstyle)|Returns or sets series gradient style of a region map chart. Read/Write.|
||[invertColor](/javascript/api/excel/excel.chartseriesloadoptions#invertcolor)|Returns or sets the fill color for negative data points in a series. Read/Write.|
||[mapOptions](/javascript/api/excel/excel.chartseriesloadoptions#mapoptions)|Encapsulates the options for a region map chart.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesloadoptions#parentlabelstrategy)|Returns or sets the series parent label strategy area for a treemap chart. Read/Write.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesloadoptions#showconnectorlines)|Specifies whether or not connector lines are shown in waterfall charts. Read/Write.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesloadoptions#showleaderlines)|Specifies whether or not leader lines are displayed for each data label in the series. Read/Write.|
||[splitValue](/javascript/api/excel/excel.chartseriesloadoptions#splitvalue)|Returns or sets the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/Write.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#xerrorbars)|Represents the error bar object of a chart series.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#yerrorbars)|Represents the error bar object of a chart series.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[binOptions](/javascript/api/excel/excel.chartseriesupdatedata#binoptions)|Encapsulates the bin options for histogram charts and pareto charts.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesupdatedata#boxwhiskeroptions)|Encapsulates the options for the box and whisker charts.|
||[bubbleScale](/javascript/api/excel/excel.chartseriesupdatedata#bubblescale)|This can be an integer value from 0 (zero) to 300, representing the percentage of the default size. This property only applies to bubble charts. Read/Write.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumcolor)|Returns or sets the color for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumtype)|Returns or sets the type for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumvalue)|Returns or sets the maximum value of a region map chart series. Read/Write.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointcolor)|Returns or sets the color for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointtype)|Returns or sets the type for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointvalue)|Returns or sets the midpoint value of a region map chart series. Read/Write.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumcolor)|Returns or sets the color for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumtype)|Returns or sets the type for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumvalue)|Returns or sets the minimum value of a region map chart series. Read/Write.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesupdatedata#gradientstyle)|Returns or sets series gradient style of a region map chart. Read/Write.|
||[invertColor](/javascript/api/excel/excel.chartseriesupdatedata#invertcolor)|Returns or sets the fill color for negative data points in a series. Read/Write.|
||[mapOptions](/javascript/api/excel/excel.chartseriesupdatedata#mapoptions)|Encapsulates the options for a region map chart.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesupdatedata#parentlabelstrategy)|Returns or sets the series parent label strategy area for a treemap chart. Read/Write.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesupdatedata#showconnectorlines)|Specifies whether or not connector lines are shown in waterfall charts. Read/Write.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesupdatedata#showleaderlines)|Specifies whether or not leader lines are displayed for each data label in the series. Read/Write.|
||[splitValue](/javascript/api/excel/excel.chartseriesupdatedata#splitvalue)|Returns or sets the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/Write.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#xerrorbars)|Represents the error bar object of a chart series.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#yerrorbars)|Represents the error bar object of a chart series.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[pivotOptions](/javascript/api/excel/excel.chartupdatedata#pivotoptions)|Encapsulates the options for a pivot chart.|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Represents the `addressLocal` property.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Represents the `columnIndex` property.|
|[ColumnPropertiesLoadOptions](/javascript/api/excel/excel.columnpropertiesloadoptions)|[columnHidden](/javascript/api/excel/excel.columnpropertiesloadoptions#columnhidden)|Specifies whether to load on the `columnHidden` property.|
||[columnIndex](/javascript/api/excel/excel.columnpropertiesloadoptions#columnindex)|Specifies whether to load on the `columnIndex` property.|
||[columnWidth](/javascript/api/excel/excel.columnpropertiesloadoptions#columnwidth)||
||[format: Excel.CellPropertiesFormatLoadOptions & {
            columnWidth?](/javascript/api/excel/excel.columnpropertiesloadoptions#format)|Specifies whether to load on the `format` property.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Returns the RangeAreas, comprising one or more rectangular ranges, the conditonal format is applied to. Read-only.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will throw an ItemNotFound error.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will return null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|The property used by the filter to do rich filter on richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Returns the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Returns the Shape object for the geometric shape. Read-only.|
|[GeometricShapeData](/javascript/api/excel/excel.geometricshapedata)|[id](/javascript/api/excel/excel.geometricshapedata#id)|Returns the shape identifier. Read-only.|
|[GeometricShapeLoadOptions](/javascript/api/excel/excel.geometricshapeloadoptions)|[$all](/javascript/api/excel/excel.geometricshapeloadoptions#$all)||
||[id](/javascript/api/excel/excel.geometricshapeloadoptions#id)|Returns the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.geometricshapeloadoptions#shape)|Returns the Shape object for the geometric shape.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Returns the number of shapes in the shape group. Read-only.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Gets a shape using its Name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Gets a shape based on its position in the collection.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Gets the loaded child items in this collection.|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[$all](/javascript/api/excel/excel.groupshapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttextdescription)|For EACH ITEM in the collection: Returns or sets the alternative description text for a Shape object.|
||[altTextTitle](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttexttitle)|For EACH ITEM in the collection: Returns or sets the alternative title text for a Shape object.|
||[connectionSiteCount](/javascript/api/excel/excel.groupshapecollectionloadoptions#connectionsitecount)|For EACH ITEM in the collection: Returns the number of connection sites on this shape. Read-only.|
||[fill](/javascript/api/excel/excel.groupshapecollectionloadoptions#fill)|For EACH ITEM in the collection: Returns the fill formatting of this shape.|
||[geometricShape](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshape)|For EACH ITEM in the collection: Returns the geometric shape associated with the shape. An error will be thrown if the shape type is not "GeometricShape".|
||[geometricShapeType](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshapetype)|For EACH ITEM in the collection: Represents the geometric shape type of this geometric shape. See Excel.GeometricShapeType for details. Returns null if the shape type is not "GeometricShape".|
||[group](/javascript/api/excel/excel.groupshapecollectionloadoptions#group)|For EACH ITEM in the collection: Returns the shape group associated with the shape. An error will be thrown if the shape type is not "GroupShape".|
||[height](/javascript/api/excel/excel.groupshapecollectionloadoptions#height)|For EACH ITEM in the collection: Represents the height, in points, of the shape.|
||[id](/javascript/api/excel/excel.groupshapecollectionloadoptions#id)|For EACH ITEM in the collection: Represents the shape identifier. Read-only.|
||[image](/javascript/api/excel/excel.groupshapecollectionloadoptions#image)|For EACH ITEM in the collection: Returns the image associated with the shape. An error will be thrown if the shape type is not "Image".|
||[left](/javascript/api/excel/excel.groupshapecollectionloadoptions#left)|For EACH ITEM in the collection: The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[level](/javascript/api/excel/excel.groupshapecollectionloadoptions#level)|For EACH ITEM in the collection: Represents the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.|
||[line](/javascript/api/excel/excel.groupshapecollectionloadoptions#line)|For EACH ITEM in the collection: Returns the line associated with the shape. An error will be thrown if the shape type is not "Line".|
||[lineFormat](/javascript/api/excel/excel.groupshapecollectionloadoptions#lineformat)|For EACH ITEM in the collection: Returns the line formatting of this shape.|
||[lockAspectRatio](/javascript/api/excel/excel.groupshapecollectionloadoptions#lockaspectratio)|For EACH ITEM in the collection: Specifies whether or not the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.groupshapecollectionloadoptions#name)|For EACH ITEM in the collection: Represents the name of the shape.|
||[parentGroup](/javascript/api/excel/excel.groupshapecollectionloadoptions#parentgroup)|For EACH ITEM in the collection: Represents the parent group of this shape.|
||[rotation](/javascript/api/excel/excel.groupshapecollectionloadoptions#rotation)|For EACH ITEM in the collection: Represents the rotation, in degrees, of the shape.|
||[textFrame](/javascript/api/excel/excel.groupshapecollectionloadoptions#textframe)|For EACH ITEM in the collection: Returns the text frame object of this shape. Read only.|
||[top](/javascript/api/excel/excel.groupshapecollectionloadoptions#top)|For EACH ITEM in the collection: The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[type](/javascript/api/excel/excel.groupshapecollectionloadoptions#type)|For EACH ITEM in the collection: Returns the type of this shape. See Excel.ShapeType for details. Read-only.|
||[visible](/javascript/api/excel/excel.groupshapecollectionloadoptions#visible)|For EACH ITEM in the collection: Represents the visibility of this shape.|
||[width](/javascript/api/excel/excel.groupshapecollectionloadoptions#width)|For EACH ITEM in the collection: Represents the width, in points, of the shape.|
||[zOrderPosition](/javascript/api/excel/excel.groupshapecollectionloadoptions#zorderposition)|For EACH ITEM in the collection: Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack. Read-only.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Gets or sets the center footer of the worksheet.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Gets or sets the center header of the worksheet.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Gets or sets the left footer of the worksheet.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Gets or sets the left header of the worksheet.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Gets or sets the right footer of the worksheet.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Gets or sets the right header of the worksheet.|
||[set(properties: Excel.HeaderFooter)](/javascript/api/excel/excel.headerfooter#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.HeaderFooterUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.headerfooter#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[HeaderFooterData](/javascript/api/excel/excel.headerfooterdata)|[centerFooter](/javascript/api/excel/excel.headerfooterdata#centerfooter)|Gets or sets the center footer of the worksheet.|
||[centerHeader](/javascript/api/excel/excel.headerfooterdata#centerheader)|Gets or sets the center header of the worksheet.|
||[leftFooter](/javascript/api/excel/excel.headerfooterdata#leftfooter)|Gets or sets the left footer of the worksheet.|
||[leftHeader](/javascript/api/excel/excel.headerfooterdata#leftheader)|Gets or sets the left header of the worksheet.|
||[rightFooter](/javascript/api/excel/excel.headerfooterdata#rightfooter)|Gets or sets the right footer of the worksheet.|
||[rightHeader](/javascript/api/excel/excel.headerfooterdata#rightheader)|Gets or sets the right header of the worksheet.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|The general header/footer, used for all pages unless even/odd or first page is specified.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|The first page header/footer, for all other pages general or even/odd is used.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|The header/footer to use for odd pages, even header/footer needs to be specified for even pages.|
||[set(properties: Excel.HeaderFooterGroup)](/javascript/api/excel/excel.headerfootergroup#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.HeaderFooterGroupUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.headerfootergroup#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Gets or sets the state of which headers/footers are set. See Excel.HeaderFooterState for details.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.|
|[HeaderFooterGroupData](/javascript/api/excel/excel.headerfootergroupdata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupdata#defaultforallpages)|The general header/footer, used for all pages unless even/odd or first page is specified.|
||[evenPages](/javascript/api/excel/excel.headerfootergroupdata#evenpages)|The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.|
||[firstPage](/javascript/api/excel/excel.headerfootergroupdata#firstpage)|The first page header/footer, for all other pages general or even/odd is used.|
||[oddPages](/javascript/api/excel/excel.headerfootergroupdata#oddpages)|The header/footer to use for odd pages, even header/footer needs to be specified for even pages.|
||[state](/javascript/api/excel/excel.headerfootergroupdata#state)|Gets or sets the state of which headers/footers are set. See Excel.HeaderFooterState for details.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupdata#usesheetmargins)|Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupdata#usesheetscale)|Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.|
|[HeaderFooterGroupLoadOptions](/javascript/api/excel/excel.headerfootergrouploadoptions)|[$all](/javascript/api/excel/excel.headerfootergrouploadoptions#$all)||
||[defaultForAllPages](/javascript/api/excel/excel.headerfootergrouploadoptions#defaultforallpages)|The general header/footer, used for all pages unless even/odd or first page is specified.|
||[evenPages](/javascript/api/excel/excel.headerfootergrouploadoptions#evenpages)|The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.|
||[firstPage](/javascript/api/excel/excel.headerfootergrouploadoptions#firstpage)|The first page header/footer, for all other pages general or even/odd is used.|
||[oddPages](/javascript/api/excel/excel.headerfootergrouploadoptions#oddpages)|The header/footer to use for odd pages, even header/footer needs to be specified for even pages.|
||[state](/javascript/api/excel/excel.headerfootergrouploadoptions#state)|Gets or sets the state of which headers/footers are set. See Excel.HeaderFooterState for details.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetmargins)|Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetscale)|Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.|
|[HeaderFooterGroupUpdateData](/javascript/api/excel/excel.headerfootergroupupdatedata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupupdatedata#defaultforallpages)|The general header/footer, used for all pages unless even/odd or first page is specified.|
||[evenPages](/javascript/api/excel/excel.headerfootergroupupdatedata#evenpages)|The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.|
||[firstPage](/javascript/api/excel/excel.headerfootergroupupdatedata#firstpage)|The first page header/footer, for all other pages general or even/odd is used.|
||[oddPages](/javascript/api/excel/excel.headerfootergroupupdatedata#oddpages)|The header/footer to use for odd pages, even header/footer needs to be specified for even pages.|
||[state](/javascript/api/excel/excel.headerfootergroupupdatedata#state)|Gets or sets the state of which headers/footers are set. See Excel.HeaderFooterState for details.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetmargins)|Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetscale)|Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.|
|[HeaderFooterLoadOptions](/javascript/api/excel/excel.headerfooterloadoptions)|[$all](/javascript/api/excel/excel.headerfooterloadoptions#$all)||
||[centerFooter](/javascript/api/excel/excel.headerfooterloadoptions#centerfooter)|Gets or sets the center footer of the worksheet.|
||[centerHeader](/javascript/api/excel/excel.headerfooterloadoptions#centerheader)|Gets or sets the center header of the worksheet.|
||[leftFooter](/javascript/api/excel/excel.headerfooterloadoptions#leftfooter)|Gets or sets the left footer of the worksheet.|
||[leftHeader](/javascript/api/excel/excel.headerfooterloadoptions#leftheader)|Gets or sets the left header of the worksheet.|
||[rightFooter](/javascript/api/excel/excel.headerfooterloadoptions#rightfooter)|Gets or sets the right footer of the worksheet.|
||[rightHeader](/javascript/api/excel/excel.headerfooterloadoptions#rightheader)|Gets or sets the right header of the worksheet.|
|[HeaderFooterUpdateData](/javascript/api/excel/excel.headerfooterupdatedata)|[centerFooter](/javascript/api/excel/excel.headerfooterupdatedata#centerfooter)|Gets or sets the center footer of the worksheet.|
||[centerHeader](/javascript/api/excel/excel.headerfooterupdatedata#centerheader)|Gets or sets the center header of the worksheet.|
||[leftFooter](/javascript/api/excel/excel.headerfooterupdatedata#leftfooter)|Gets or sets the left footer of the worksheet.|
||[leftHeader](/javascript/api/excel/excel.headerfooterupdatedata#leftheader)|Gets or sets the left header of the worksheet.|
||[rightFooter](/javascript/api/excel/excel.headerfooterupdatedata#rightfooter)|Gets or sets the right footer of the worksheet.|
||[rightHeader](/javascript/api/excel/excel.headerfooterupdatedata#rightheader)|Gets or sets the right header of the worksheet.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Returns the format of the image. Read-only.|
||[id](/javascript/api/excel/excel.image#id)|Represents the shape identifier for the image object. Read-only.|
||[shape](/javascript/api/excel/excel.image#shape)|Returns the Shape object associated with the image. Read-only.|
|[ImageData](/javascript/api/excel/excel.imagedata)|[format](/javascript/api/excel/excel.imagedata#format)|Returns the format of the image. Read-only.|
||[id](/javascript/api/excel/excel.imagedata#id)|Represents the shape identifier for the image object. Read-only.|
|[ImageLoadOptions](/javascript/api/excel/excel.imageloadoptions)|[$all](/javascript/api/excel/excel.imageloadoptions#$all)||
||[format](/javascript/api/excel/excel.imageloadoptions#format)|Returns the format of the image. Read-only.|
||[id](/javascript/api/excel/excel.imageloadoptions#id)|Represents the shape identifier for the image object. Read-only.|
||[shape](/javascript/api/excel/excel.imageloadoptions#shape)|Returns the Shape object associated with the image.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.|
||[set(properties: Excel.IterativeCalculation)](/javascript/api/excel/excel.iterativecalculation#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.IterativeCalculationUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.iterativecalculation#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[IterativeCalculationData](/javascript/api/excel/excel.iterativecalculationdata)|[enabled](/javascript/api/excel/excel.iterativecalculationdata#enabled)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationdata#maxchange)|Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationdata#maxiteration)|Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.|
|[IterativeCalculationLoadOptions](/javascript/api/excel/excel.iterativecalculationloadoptions)|[$all](/javascript/api/excel/excel.iterativecalculationloadoptions#$all)||
||[enabled](/javascript/api/excel/excel.iterativecalculationloadoptions#enabled)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationloadoptions#maxchange)|Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationloadoptions#maxiteration)|Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.|
|[IterativeCalculationUpdateData](/javascript/api/excel/excel.iterativecalculationupdatedata)|[enabled](/javascript/api/excel/excel.iterativecalculationupdatedata#enabled)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationupdatedata#maxchange)|Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationupdatedata#maxiteration)|Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.|
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
||[set(properties: Excel.Line)](/javascript/api/excel/excel.line#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.LineUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.line#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[LineData](/javascript/api/excel/excel.linedata)|[beginArrowheadLength](/javascript/api/excel/excel.linedata#beginarrowheadlength)|Represents the length of the arrowhead at the beginning of the specified line.|
||[beginArrowheadStyle](/javascript/api/excel/excel.linedata#beginarrowheadstyle)|Represents the style of the arrowhead at the beginning of the specified line.|
||[beginArrowheadWidth](/javascript/api/excel/excel.linedata#beginarrowheadwidth)|Represents the width of the arrowhead at the beginning of the specified line.|
||[beginConnectedSite](/javascript/api/excel/excel.linedata#beginconnectedsite)|Represents the connection site to which the beginning of a connector is connected. Read-only. Returns null when the beginning of the line is not attached to any shape.|
||[connectorType](/javascript/api/excel/excel.linedata#connectortype)|Represents the connector type for the line.|
||[endArrowheadLength](/javascript/api/excel/excel.linedata#endarrowheadlength)|Represents the length of the arrowhead at the end of the specified line.|
||[endArrowheadStyle](/javascript/api/excel/excel.linedata#endarrowheadstyle)|Represents the style of the arrowhead at the end of the specified line.|
||[endArrowheadWidth](/javascript/api/excel/excel.linedata#endarrowheadwidth)|Represents the width of the arrowhead at the end of the specified line.|
||[endConnectedSite](/javascript/api/excel/excel.linedata#endconnectedsite)|Represents the connection site to which the end of a connector is connected. Read-only. Returns null when the end of the line is not attached to any shape.|
||[id](/javascript/api/excel/excel.linedata#id)|Represents the shape identifier. Read-only.|
||[isBeginConnected](/javascript/api/excel/excel.linedata#isbeginconnected)|Specifies whether or not the beginning of the specified line is connected to a shape. Read-only.|
||[isEndConnected](/javascript/api/excel/excel.linedata#isendconnected)|Specifies whether or not the end of the specified line is connected to a shape. Read-only.|
|[LineLoadOptions](/javascript/api/excel/excel.lineloadoptions)|[$all](/javascript/api/excel/excel.lineloadoptions#$all)||
||[beginArrowheadLength](/javascript/api/excel/excel.lineloadoptions#beginarrowheadlength)|Represents the length of the arrowhead at the beginning of the specified line.|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#beginarrowheadstyle)|Represents the style of the arrowhead at the beginning of the specified line.|
||[beginArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#beginarrowheadwidth)|Represents the width of the arrowhead at the beginning of the specified line.|
||[beginConnectedShape](/javascript/api/excel/excel.lineloadoptions#beginconnectedshape)|Represents the shape to which the beginning of the specified line is attached.|
||[beginConnectedSite](/javascript/api/excel/excel.lineloadoptions#beginconnectedsite)|Represents the connection site to which the beginning of a connector is connected. Read-only. Returns null when the beginning of the line is not attached to any shape.|
||[connectorType](/javascript/api/excel/excel.lineloadoptions#connectortype)|Represents the connector type for the line.|
||[endArrowheadLength](/javascript/api/excel/excel.lineloadoptions#endarrowheadlength)|Represents the length of the arrowhead at the end of the specified line.|
||[endArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#endarrowheadstyle)|Represents the style of the arrowhead at the end of the specified line.|
||[endArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#endarrowheadwidth)|Represents the width of the arrowhead at the end of the specified line.|
||[endConnectedShape](/javascript/api/excel/excel.lineloadoptions#endconnectedshape)|Represents the shape to which the end of the specified line is attached.|
||[endConnectedSite](/javascript/api/excel/excel.lineloadoptions#endconnectedsite)|Represents the connection site to which the end of a connector is connected. Read-only. Returns null when the end of the line is not attached to any shape.|
||[id](/javascript/api/excel/excel.lineloadoptions#id)|Represents the shape identifier. Read-only.|
||[isBeginConnected](/javascript/api/excel/excel.lineloadoptions#isbeginconnected)|Specifies whether or not the beginning of the specified line is connected to a shape. Read-only.|
||[isEndConnected](/javascript/api/excel/excel.lineloadoptions#isendconnected)|Specifies whether or not the end of the specified line is connected to a shape. Read-only.|
||[shape](/javascript/api/excel/excel.lineloadoptions#shape)|Returns the Shape object associated with the line.|
|[LineUpdateData](/javascript/api/excel/excel.lineupdatedata)|[beginArrowheadLength](/javascript/api/excel/excel.lineupdatedata#beginarrowheadlength)|Represents the length of the arrowhead at the beginning of the specified line.|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#beginarrowheadstyle)|Represents the style of the arrowhead at the beginning of the specified line.|
||[beginArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#beginarrowheadwidth)|Represents the width of the arrowhead at the beginning of the specified line.|
||[connectorType](/javascript/api/excel/excel.lineupdatedata#connectortype)|Represents the connector type for the line.|
||[endArrowheadLength](/javascript/api/excel/excel.lineupdatedata#endarrowheadlength)|Represents the length of the arrowhead at the end of the specified line.|
||[endArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#endarrowheadstyle)|Represents the style of the arrowhead at the end of the specified line.|
||[endArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#endarrowheadwidth)|Represents the width of the arrowhead at the end of the specified line.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Deletes a page break object.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Gets the first cell after the page break.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Represents the column index for the page break|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Represents the row index for the page break|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Adds a page break before the top-left cell of the range specified.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Gets the number of page breaks in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Gets a page break object via the index.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Gets the loaded child items in this collection.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Resets all manual page breaks in the collection.|
|[PageBreakCollectionLoadOptions](/javascript/api/excel/excel.pagebreakcollectionloadoptions)|[$all](/javascript/api/excel/excel.pagebreakcollectionloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#columnindex)|For EACH ITEM in the collection: Represents the column index for the page break|
||[rowIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#rowindex)|For EACH ITEM in the collection: Represents the row index for the page break|
|[PageBreakData](/javascript/api/excel/excel.pagebreakdata)|[columnIndex](/javascript/api/excel/excel.pagebreakdata#columnindex)|Represents the column index for the page break|
||[rowIndex](/javascript/api/excel/excel.pagebreakdata#rowindex)|Represents the row index for the page break|
|[PageBreakLoadOptions](/javascript/api/excel/excel.pagebreakloadoptions)|[$all](/javascript/api/excel/excel.pagebreakloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakloadoptions#columnindex)|Represents the column index for the page break|
||[rowIndex](/javascript/api/excel/excel.pagebreakloadoptions#rowindex)|Represents the row index for the page break|
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
||[set(properties: Excel.PageLayout)](/javascript/api/excel/excel.pagelayout#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.PageLayoutUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.pagelayout#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Sets the worksheet's print area.|
||[setPrintMargins(unit: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Sets the worksheet's page margins with units.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Sets the worksheet's page margins with units.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Gets or sets the worksheet's top margin, in points, for use when printing.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Gets or sets the worksheet's print zoom options.|
|[PageLayoutData](/javascript/api/excel/excel.pagelayoutdata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutdata#blackandwhite)|Gets or sets the worksheet's black and white print option.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutdata#bottommargin)|Gets or sets the worksheet's bottom page margin to use for printing in points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutdata#centerhorizontally)|Gets or sets the worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutdata#centervertically)|Gets or sets the worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.|
||[draftMode](/javascript/api/excel/excel.pagelayoutdata#draftmode)|Gets or sets the worksheet's draft mode option. If true the sheet will be printed without graphics.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutdata#firstpagenumber)|Gets or sets the worksheet's first page number to print. Null value represents "auto" page numbering.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutdata#footermargin)|Gets or sets the worksheet's footer margin, in points, for use when printing.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutdata#headermargin)|Gets or sets the worksheet's header margin, in points, for use when printing.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutdata#headersfooters)|Header and footer configuration for the worksheet.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutdata#leftmargin)|Gets or sets the worksheet's left margin, in points, for use when printing.|
||[orientation](/javascript/api/excel/excel.pagelayoutdata#orientation)|Gets or sets the worksheet's orientation of the page.|
||[paperSize](/javascript/api/excel/excel.pagelayoutdata#papersize)|Gets or sets the worksheet's paper size of the page.|
||[printComments](/javascript/api/excel/excel.pagelayoutdata#printcomments)|Gets or sets whether the worksheet's comments should be displayed when printing.|
||[printErrors](/javascript/api/excel/excel.pagelayoutdata#printerrors)|Gets or sets the worksheet's print errors option.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutdata#printgridlines)|Gets or sets the worksheet's print gridlines flag. This flag determines whether gridlines will be printed or not.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutdata#printheadings)|Gets or sets the worksheet's print headings flag. This flag determines whether headings will be printed or not.|
||[printOrder](/javascript/api/excel/excel.pagelayoutdata#printorder)|Gets or sets the worksheet's page print order option. This specifies the order to use for processing the page number printed.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutdata#rightmargin)|Gets or sets the worksheet's right margin, in points, for use when printing.|
||[topMargin](/javascript/api/excel/excel.pagelayoutdata#topmargin)|Gets or sets the worksheet's top margin, in points, for use when printing.|
||[zoom](/javascript/api/excel/excel.pagelayoutdata#zoom)|Gets or sets the worksheet's print zoom options.|
|[PageLayoutLoadOptions](/javascript/api/excel/excel.pagelayoutloadoptions)|[$all](/javascript/api/excel/excel.pagelayoutloadoptions#$all)||
||[blackAndWhite](/javascript/api/excel/excel.pagelayoutloadoptions#blackandwhite)|Gets or sets the worksheet's black and white print option.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutloadoptions#bottommargin)|Gets or sets the worksheet's bottom page margin to use for printing in points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutloadoptions#centerhorizontally)|Gets or sets the worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutloadoptions#centervertically)|Gets or sets the worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.|
||[draftMode](/javascript/api/excel/excel.pagelayoutloadoptions#draftmode)|Gets or sets the worksheet's draft mode option. If true the sheet will be printed without graphics.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutloadoptions#firstpagenumber)|Gets or sets the worksheet's first page number to print. Null value represents "auto" page numbering.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#footermargin)|Gets or sets the worksheet's footer margin, in points, for use when printing.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#headermargin)|Gets or sets the worksheet's header margin, in points, for use when printing.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutloadoptions#headersfooters)|Header and footer configuration for the worksheet.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutloadoptions#leftmargin)|Gets or sets the worksheet's left margin, in points, for use when printing.|
||[orientation](/javascript/api/excel/excel.pagelayoutloadoptions#orientation)|Gets or sets the worksheet's orientation of the page.|
||[paperSize](/javascript/api/excel/excel.pagelayoutloadoptions#papersize)|Gets or sets the worksheet's paper size of the page.|
||[printComments](/javascript/api/excel/excel.pagelayoutloadoptions#printcomments)|Gets or sets whether the worksheet's comments should be displayed when printing.|
||[printErrors](/javascript/api/excel/excel.pagelayoutloadoptions#printerrors)|Gets or sets the worksheet's print errors option.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutloadoptions#printgridlines)|Gets or sets the worksheet's print gridlines flag. This flag determines whether gridlines will be printed or not.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutloadoptions#printheadings)|Gets or sets the worksheet's print headings flag. This flag determines whether headings will be printed or not.|
||[printOrder](/javascript/api/excel/excel.pagelayoutloadoptions#printorder)|Gets or sets the worksheet's page print order option. This specifies the order to use for processing the page number printed.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutloadoptions#rightmargin)|Gets or sets the worksheet's right margin, in points, for use when printing.|
||[topMargin](/javascript/api/excel/excel.pagelayoutloadoptions#topmargin)|Gets or sets the worksheet's top margin, in points, for use when printing.|
||[zoom](/javascript/api/excel/excel.pagelayoutloadoptions#zoom)|Gets or sets the worksheet's print zoom options.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Represents the page layout bottom margin in the unit specified to use for printing.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Represents the page layout footer margin in the unit specified to use for printing.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Represents the page layout header margin in the unit specified to use for printing.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Represents the page layout left margin in the unit specified to use for printing.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Represents the page layout right margin in the unit specified to use for printing.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Represents the page layout top margin in the unit specified to use for printing.|
|[PageLayoutUpdateData](/javascript/api/excel/excel.pagelayoutupdatedata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutupdatedata#blackandwhite)|Gets or sets the worksheet's black and white print option.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutupdatedata#bottommargin)|Gets or sets the worksheet's bottom page margin to use for printing in points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutupdatedata#centerhorizontally)|Gets or sets the worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutupdatedata#centervertically)|Gets or sets the worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.|
||[draftMode](/javascript/api/excel/excel.pagelayoutupdatedata#draftmode)|Gets or sets the worksheet's draft mode option. If true the sheet will be printed without graphics.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutupdatedata#firstpagenumber)|Gets or sets the worksheet's first page number to print. Null value represents "auto" page numbering.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#footermargin)|Gets or sets the worksheet's footer margin, in points, for use when printing.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#headermargin)|Gets or sets the worksheet's header margin, in points, for use when printing.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutupdatedata#headersfooters)|Header and footer configuration for the worksheet.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutupdatedata#leftmargin)|Gets or sets the worksheet's left margin, in points, for use when printing.|
||[orientation](/javascript/api/excel/excel.pagelayoutupdatedata#orientation)|Gets or sets the worksheet's orientation of the page.|
||[paperSize](/javascript/api/excel/excel.pagelayoutupdatedata#papersize)|Gets or sets the worksheet's paper size of the page.|
||[printComments](/javascript/api/excel/excel.pagelayoutupdatedata#printcomments)|Gets or sets whether the worksheet's comments should be displayed when printing.|
||[printErrors](/javascript/api/excel/excel.pagelayoutupdatedata#printerrors)|Gets or sets the worksheet's print errors option.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutupdatedata#printgridlines)|Gets or sets the worksheet's print gridlines flag. This flag determines whether gridlines will be printed or not.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutupdatedata#printheadings)|Gets or sets the worksheet's print headings flag. This flag determines whether headings will be printed or not.|
||[printOrder](/javascript/api/excel/excel.pagelayoutupdatedata#printorder)|Gets or sets the worksheet's page print order option. This specifies the order to use for processing the page number printed.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutupdatedata#rightmargin)|Gets or sets the worksheet's right margin, in points, for use when printing.|
||[topMargin](/javascript/api/excel/excel.pagelayoutupdatedata#topmargin)|Gets or sets the worksheet's top margin, in points, for use when printing.|
||[zoom](/javascript/api/excel/excel.pagelayoutupdatedata#zoom)|Gets or sets the worksheet's print zoom options.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Number of pages to fit horizontally. This value can be null if percentage scale is used.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|Print page scale value can be between 10 and 400. This value can be null if fit to page tall or wide is specified.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Number of pages to fit vertically. This value can be null if percentage scale is used.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: "Ascending" \| "Descending", valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when|
||[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Specifies whether formatting will be automatically formatted when it’s refreshed or when fields are moved|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|
||[getPivotItems(axis: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Specifies whether formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: "Ascending" \| "Descending")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context. This behaves identically to applying an autosort from the UI.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context. This behaves identically to applying an autosort from the UI.|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutdata#autoformat)|Specifies whether formatting will be automatically formatted when it’s refreshed or when fields are moved|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutdata#preserveformatting)|Specifies whether formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[autoFormat](/javascript/api/excel/excel.pivotlayoutloadoptions#autoformat)|Specifies whether formatting will be automatically formatted when it’s refreshed or when fields are moved|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutloadoptions#preserveformatting)|Specifies whether formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutupdatedata#autoformat)|Specifies whether formatting will be automatically formatted when it’s refreshed or when fields are moved|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutupdatedata#preserveformatting)|Specifies whether formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Specifies whether the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Specifies whether the PivotTable uses custom lists when sorting.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottablecollectionloadoptions#enabledatavalueediting)|For EACH ITEM in the collection: Specifies whether the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottablecollectionloadoptions#usecustomsortlists)|For EACH ITEM in the collection: Specifies whether the PivotTable uses custom lists when sorting.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottabledata#enabledatavalueediting)|Specifies whether the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottabledata#usecustomsortlists)|Specifies whether the PivotTable uses custom lists when sorting.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableloadoptions#enabledatavalueediting)|Specifies whether the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableloadoptions#usecustomsortlists)|Specifies whether the PivotTable uses custom lists when sorting.|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableupdatedata#enabledatavalueediting)|Specifies whether the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableupdatedata#usecustomsortlists)|Specifies whether the PivotTable uses custom lists when sorting.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Fills range from the current range to the destination range.|
||[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Fills range from the current range to the destination range.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Converts the range cells with datatypes into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Converts the range cells into linked datatype in the worksheet.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copies cell data or formatting from the source range or RangeAreas to the current range.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copies cell data or formatting from the source range or RangeAreas to the current range.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Finds the given string based on the criteria specified.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Finds the given string based on the criteria specified.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Does FlashFill to current range.Flash Fill will automatically fills data when it senses a pattern, so the range must be single column range and have data around in order to find pattern.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Returns a 2D array, encapsulating the data for each cell's font, fill, borders, alignment, and other properties.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Returns a single-dimensional array, encapsulating the data for each column's font, fill, borders, alignment, and other properties.  For properties that are not consistent across each cell within a given column, null will be returned.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Returns a single-dimensional array, encapsulating the data for each row's font, fill, borders, alignment, and other properties.  For properties that are not consistent across each cell within a given row, null will be returned.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.|
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
||[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Converts all cells in the RangeAreas with datatypes into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Converts all cells in the RangeAreas into linked datatype.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Returns a RangeAreas object that represents the entire columns of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11, H2", it returns a RangeAreas that represents columns "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Returns a RangeAreas object that represents the entire rows of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11", it returns a RangeAreas that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, an ItemNotFound error will be thrown.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, a null object is returned.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Returns an RangeAreas object that is shifted by the specific row and column offset. The dimension of the returned RangeAreas will match the original object. If the resulting RangeAreas is forced outside the bounds of the worksheet grid, an error will be thrown.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.|
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
||[set(properties: Excel.RangeAreas)](/javascript/api/excel/excel.rangeareas#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RangeAreasUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.rangeareas#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Sets the RangeAreas to be recalculated when the next recalculation occurs.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Represents the style for all ranges in this RangeAreas object.|
||[track()](/javascript/api/excel/excel.rangeareas#track--)|Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.|
|[RangeAreasData](/javascript/api/excel/excel.rangeareasdata)|[address](/javascript/api/excel/excel.rangeareasdata#address)|Returns the RageAreas reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g. "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.|
||[addressLocal](/javascript/api/excel/excel.rangeareasdata#addresslocal)|Returns the RageAreas reference in the user locale. Read-only.|
||[areaCount](/javascript/api/excel/excel.rangeareasdata#areacount)|Returns the number of rectangular ranges that comprise this RangeAreas object.|
||[areas](/javascript/api/excel/excel.rangeareasdata#areas)|Returns a collection of rectangular ranges that comprise this RangeAreas object.|
||[cellCount](/javascript/api/excel/excel.rangeareasdata#cellcount)|Returns the number of cells in the RangeAreas object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareasdata#conditionalformats)|Returns a collection of ConditionalFormats that intersect with any cells in this RangeAreas object. Read-only.|
||[dataValidation](/javascript/api/excel/excel.rangeareasdata#datavalidation)|Returns a dataValidation object for all ranges in the RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasdata#format)|Returns a rangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object. Read-only.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasdata#isentirecolumn)|Indicates whether all the ranges on this RangeAreas object represent entire columns (e.g., "A:C, Q:Z"). Read-only.|
||[isEntireRow](/javascript/api/excel/excel.rangeareasdata#isentirerow)|Indicates whether all the ranges on this RangeAreas object represent entire rows (e.g., "1:3, 5:7"). Read-only.|
||[style](/javascript/api/excel/excel.rangeareasdata#style)|Represents the style for all ranges in this RangeAreas object.|
|[RangeAreasLoadOptions](/javascript/api/excel/excel.rangeareasloadoptions)|[$all](/javascript/api/excel/excel.rangeareasloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeareasloadoptions#address)|Returns the RageAreas reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g. "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.|
||[addressLocal](/javascript/api/excel/excel.rangeareasloadoptions#addresslocal)|Returns the RageAreas reference in the user locale. Read-only.|
||[areaCount](/javascript/api/excel/excel.rangeareasloadoptions#areacount)|Returns the number of rectangular ranges that comprise this RangeAreas object.|
||[cellCount](/javascript/api/excel/excel.rangeareasloadoptions#cellcount)|Returns the number of cells in the RangeAreas object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|
||[dataValidation](/javascript/api/excel/excel.rangeareasloadoptions#datavalidation)|Returns a dataValidation object for all ranges in the RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasloadoptions#format)|Returns a rangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasloadoptions#isentirecolumn)|Indicates whether all the ranges on this RangeAreas object represent entire columns (e.g., "A:C, Q:Z"). Read-only.|
||[isEntireRow](/javascript/api/excel/excel.rangeareasloadoptions#isentirerow)|Indicates whether all the ranges on this RangeAreas object represent entire rows (e.g., "1:3, 5:7"). Read-only.|
||[style](/javascript/api/excel/excel.rangeareasloadoptions#style)|Represents the style for all ranges in this RangeAreas object.|
||[worksheet](/javascript/api/excel/excel.rangeareasloadoptions#worksheet)|Returns the worksheet for the current RangeAreas.|
|[RangeAreasUpdateData](/javascript/api/excel/excel.rangeareasupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeareasupdatedata#datavalidation)|Returns a dataValidation object for all ranges in the RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasupdatedata#format)|Returns a rangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object.|
||[style](/javascript/api/excel/excel.rangeareasupdatedata#style)|Represents the style for all ranges in this RangeAreas object.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Borders, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderCollectionLoadOptions](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionloadoptions#tintandshade)|For EACH ITEM in the collection: Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderCollectionUpdateData](/javascript/api/excel/excel.rangebordercollectionupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionupdatedata#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Borders, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderData](/javascript/api/excel/excel.rangeborderdata)|[tintAndShade](/javascript/api/excel/excel.rangeborderdata#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderLoadOptions](/javascript/api/excel/excel.rangeborderloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangeborderloadoptions#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderUpdateData](/javascript/api/excel/excel.rangeborderupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangeborderupdatedata#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Returns the number of ranges in the RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Returns the range object based on its position in the RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Gets the loaded child items in this collection.|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[$all](/javascript/api/excel/excel.rangecollectionloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangecollectionloadoptions#address)|For EACH ITEM in the collection: Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. "Sheet1!A1:B4"). Read-only.|
||[addressLocal](/javascript/api/excel/excel.rangecollectionloadoptions#addresslocal)|For EACH ITEM in the collection: Represents range reference for the specified range in the language of the user. Read-only.|
||[cellCount](/javascript/api/excel/excel.rangecollectionloadoptions#cellcount)|For EACH ITEM in the collection: Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|
||[columnCount](/javascript/api/excel/excel.rangecollectionloadoptions#columncount)|For EACH ITEM in the collection: Represents the total number of columns in the range. Read-only.|
||[columnHidden](/javascript/api/excel/excel.rangecollectionloadoptions#columnhidden)|For EACH ITEM in the collection: Represents if all columns of the current range are hidden.|
||[columnIndex](/javascript/api/excel/excel.rangecollectionloadoptions#columnindex)|For EACH ITEM in the collection: Represents the column number of the first cell in the range. Zero-indexed. Read-only.|
||[dataValidation](/javascript/api/excel/excel.rangecollectionloadoptions#datavalidation)|For EACH ITEM in the collection: Returns a data validation object.|
||[format](/javascript/api/excel/excel.rangecollectionloadoptions#format)|For EACH ITEM in the collection: Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.|
||[formulas](/javascript/api/excel/excel.rangecollectionloadoptions#formulas)|For EACH ITEM in the collection: Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangecollectionloadoptions#formulaslocal)|For EACH ITEM in the collection: Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[formulasR1C1](/javascript/api/excel/excel.rangecollectionloadoptions#formulasr1c1)|For EACH ITEM in the collection: Represents the formula in R1C1-style notation.|
||[hidden](/javascript/api/excel/excel.rangecollectionloadoptions#hidden)|For EACH ITEM in the collection: Represents if all cells of the current range are hidden. Read-only.|
||[hyperlink](/javascript/api/excel/excel.rangecollectionloadoptions#hyperlink)|For EACH ITEM in the collection: Represents the hyperlink for the current range.|
||[isEntireColumn](/javascript/api/excel/excel.rangecollectionloadoptions#isentirecolumn)|For EACH ITEM in the collection: Represents if the current range is an entire column. Read-only.|
||[isEntireRow](/javascript/api/excel/excel.rangecollectionloadoptions#isentirerow)|For EACH ITEM in the collection: Represents if the current range is an entire row. Read-only.|
||[linkedDataTypeState](/javascript/api/excel/excel.rangecollectionloadoptions#linkeddatatypestate)|For EACH ITEM in the collection: Represents the data type state of each cell. Read-only.|
||[numberFormat](/javascript/api/excel/excel.rangecollectionloadoptions#numberformat)|For EACH ITEM in the collection: Represents Excel's number format code for the given range.|
||[numberFormatLocal](/javascript/api/excel/excel.rangecollectionloadoptions#numberformatlocal)|For EACH ITEM in the collection: Represents Excel's number format code for the given range as a string in the language of the user.|
||[rowCount](/javascript/api/excel/excel.rangecollectionloadoptions#rowcount)|For EACH ITEM in the collection: Returns the total number of rows in the range. Read-only.|
||[rowHidden](/javascript/api/excel/excel.rangecollectionloadoptions#rowhidden)|For EACH ITEM in the collection: Represents if all rows of the current range are hidden.|
||[rowIndex](/javascript/api/excel/excel.rangecollectionloadoptions#rowindex)|For EACH ITEM in the collection: Returns the row number of the first cell in the range. Zero-indexed. Read-only.|
||[style](/javascript/api/excel/excel.rangecollectionloadoptions#style)|For EACH ITEM in the collection: Represents the style of the current range.|
||[text](/javascript/api/excel/excel.rangecollectionloadoptions#text)|For EACH ITEM in the collection: Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.rangecollectionloadoptions#valuetypes)|For EACH ITEM in the collection: Represents the type of data of each cell. Read-only.|
||[values](/javascript/api/excel/excel.rangecollectionloadoptions#values)|For EACH ITEM in the collection: Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
||[worksheet](/javascript/api/excel/excel.rangecollectionloadoptions#worksheet)|For EACH ITEM in the collection: The worksheet containing the current range.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[linkedDataTypeState](/javascript/api/excel/excel.rangedata#linkeddatatypestate)|Represents the data type state of each cell. Read-only.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Gets or sets the pattern of a Range. See Excel.FillPattern for details. LinearGradient and RectangularGradient are not supported.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Sets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Returns or sets a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFillData](/javascript/api/excel/excel.rangefilldata)|[pattern](/javascript/api/excel/excel.rangefilldata#pattern)|Gets or sets the pattern of a Range. See Excel.FillPattern for details. LinearGradient and RectangularGradient are not supported.|
||[patternColor](/javascript/api/excel/excel.rangefilldata#patterncolor)|Sets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefilldata#patterntintandshade)|Returns or sets a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
||[tintAndShade](/javascript/api/excel/excel.rangefilldata#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFillLoadOptions](/javascript/api/excel/excel.rangefillloadoptions)|[pattern](/javascript/api/excel/excel.rangefillloadoptions#pattern)|Gets or sets the pattern of a Range. See Excel.FillPattern for details. LinearGradient and RectangularGradient are not supported.|
||[patternColor](/javascript/api/excel/excel.rangefillloadoptions#patterncolor)|Sets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillloadoptions#patterntintandshade)|Returns or sets a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
||[tintAndShade](/javascript/api/excel/excel.rangefillloadoptions#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFillUpdateData](/javascript/api/excel/excel.rangefillupdatedata)|[pattern](/javascript/api/excel/excel.rangefillupdatedata#pattern)|Gets or sets the pattern of a Range. See Excel.FillPattern for details. LinearGradient and RectangularGradient are not supported.|
||[patternColor](/javascript/api/excel/excel.rangefillupdatedata#patterncolor)|Sets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillupdatedata#patterntintandshade)|Returns or sets a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
||[tintAndShade](/javascript/api/excel/excel.rangefillupdatedata#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Represents the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Represents the Subscript status of font.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Represents the Superscript status of font.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFontData](/javascript/api/excel/excel.rangefontdata)|[strikethrough](/javascript/api/excel/excel.rangefontdata#strikethrough)|Represents the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.|
||[subscript](/javascript/api/excel/excel.rangefontdata#subscript)|Represents the Subscript status of font.|
||[superscript](/javascript/api/excel/excel.rangefontdata#superscript)|Represents the Superscript status of font.|
||[tintAndShade](/javascript/api/excel/excel.rangefontdata#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFontLoadOptions](/javascript/api/excel/excel.rangefontloadoptions)|[strikethrough](/javascript/api/excel/excel.rangefontloadoptions#strikethrough)|Represents the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.|
||[subscript](/javascript/api/excel/excel.rangefontloadoptions#subscript)|Represents the Subscript status of font.|
||[superscript](/javascript/api/excel/excel.rangefontloadoptions#superscript)|Represents the Superscript status of font.|
||[tintAndShade](/javascript/api/excel/excel.rangefontloadoptions#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFontUpdateData](/javascript/api/excel/excel.rangefontupdatedata)|[strikethrough](/javascript/api/excel/excel.rangefontupdatedata#strikethrough)|Represents the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.|
||[subscript](/javascript/api/excel/excel.rangefontupdatedata#subscript)|Represents the Subscript status of font.|
||[superscript](/javascript/api/excel/excel.rangefontupdatedata#superscript)|Represents the Superscript status of font.|
||[tintAndShade](/javascript/api/excel/excel.rangefontupdatedata#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Indicates if text is automatically indented when text alignment is set to equal distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|An integer from 0 to 250 that indicates the indent level.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|The reading order for the range.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[autoIndent](/javascript/api/excel/excel.rangeformatdata#autoindent)|Indicates if text is automatically indented when text alignment is set to equal distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformatdata#indentlevel)|An integer from 0 to 250 that indicates the indent level.|
||[readingOrder](/javascript/api/excel/excel.rangeformatdata#readingorder)|The reading order for the range.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatdata#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[autoIndent](/javascript/api/excel/excel.rangeformatloadoptions#autoindent)|Indicates if text is automatically indented when text alignment is set to equal distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformatloadoptions#indentlevel)|An integer from 0 to 250 that indicates the indent level.|
||[readingOrder](/javascript/api/excel/excel.rangeformatloadoptions#readingorder)|The reading order for the range.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatloadoptions#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[autoIndent](/javascript/api/excel/excel.rangeformatupdatedata#autoindent)|Indicates if text is automatically indented when text alignment is set to equal distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformatupdatedata#indentlevel)|An integer from 0 to 250 that indicates the indent level.|
||[readingOrder](/javascript/api/excel/excel.rangeformatupdatedata#readingorder)|The reading order for the range.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatupdatedata#shrinktofit)|Indicates if text automatically shrinks to fit in the available column width.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[linkedDataTypeState](/javascript/api/excel/excel.rangeloadoptions#linkeddatatypestate)|Represents the data type state of each cell. Read-only.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Number of duplicated rows removed by the operation.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Number of remaining unique rows present in the resulting range.|
|[RemoveDuplicatesResultData](/javascript/api/excel/excel.removeduplicatesresultdata)|[removed](/javascript/api/excel/excel.removeduplicatesresultdata#removed)|Number of duplicated rows removed by the operation.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultdata#uniqueremaining)|Number of remaining unique rows present in the resulting range.|
|[RemoveDuplicatesResultLoadOptions](/javascript/api/excel/excel.removeduplicatesresultloadoptions)|[$all](/javascript/api/excel/excel.removeduplicatesresultloadoptions#$all)||
||[removed](/javascript/api/excel/excel.removeduplicatesresultloadoptions#removed)|Number of duplicated rows removed by the operation.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultloadoptions#uniqueremaining)|Number of remaining unique rows present in the resulting range.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Specifies whether the match needs to be complete or partial. Default is false (partial).|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Specifies whether the match is case sensitive. Default is false (insensitive).|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Represents the `address` property.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Represents the `addressLocal` property.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Represents the `rowIndex` property.|
|[RowPropertiesLoadOptions](/javascript/api/excel/excel.rowpropertiesloadoptions)|[format: Excel.CellPropertiesFormatLoadOptions & {
            rowHeight?](/javascript/api/excel/excel.rowpropertiesloadoptions#format)|Specifies whether to load on the `format` property.|
||[rowHeight](/javascript/api/excel/excel.rowpropertiesloadoptions#rowheight)||
||[rowHidden](/javascript/api/excel/excel.rowpropertiesloadoptions#rowhidden)|Specifies whether to load on the `rowHidden` property.|
||[rowIndex](/javascript/api/excel/excel.rowpropertiesloadoptions#rowindex)|Specifies whether to load on the `rowIndex` property.|
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
||[getAsImage(format: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#getasimage-format-)|Converts the shape to an image and returns the image as a base64-encoded string. The DPI is 96. The only supported formats are `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, and `Excel.PictureFormat.GIF`.|
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
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.|
||[set(properties: Excel.Shape)](/javascript/api/excel/excel.shape#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ShapeUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.shape#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[setZOrder(position: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-position-)|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[top](/javascript/api/excel/excel.shape#top)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[visible](/javascript/api/excel/excel.shape#visible)|Represents the visibility of this shape.|
||[width](/javascript/api/excel/excel.shape#width)|Represents the width, in points, of the shape.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Gets the id of the activated shape.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Gets the id of the worksheet in which the shape is activated.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus")](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Adds a geometric shape to the worksheet. Returns a Shape object that represents the new shape.|
||[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Adds a geometric shape to the worksheet. Returns a Shape object that represents the new shape.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Groups a subset of shapes in this collection's worksheet. Returns a Shape object that represents the new group of shapes.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Creates an image from a base64-encoded string and adds it to the worksheet. Returns the Shape object that represents the new image.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adds a line to worksheet. Returns a Shape object that represents the new line.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adds a line to worksheet. Returns a Shape object that represents the new line.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Adds a text box to the worksheet with the provided text as the content. Returns a Shape object that represents the new text box.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Returns the number of shapes in the worksheet. Read-only.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Gets a shape using its Name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Gets a shape using its position in the collection.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Gets the loaded child items in this collection.|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[$all](/javascript/api/excel/excel.shapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapecollectionloadoptions#alttextdescription)|For EACH ITEM in the collection: Returns or sets the alternative description text for a Shape object.|
||[altTextTitle](/javascript/api/excel/excel.shapecollectionloadoptions#alttexttitle)|For EACH ITEM in the collection: Returns or sets the alternative title text for a Shape object.|
||[connectionSiteCount](/javascript/api/excel/excel.shapecollectionloadoptions#connectionsitecount)|For EACH ITEM in the collection: Returns the number of connection sites on this shape. Read-only.|
||[fill](/javascript/api/excel/excel.shapecollectionloadoptions#fill)|For EACH ITEM in the collection: Returns the fill formatting of this shape.|
||[geometricShape](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshape)|For EACH ITEM in the collection: Returns the geometric shape associated with the shape. An error will be thrown if the shape type is not "GeometricShape".|
||[geometricShapeType](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshapetype)|For EACH ITEM in the collection: Represents the geometric shape type of this geometric shape. See Excel.GeometricShapeType for details. Returns null if the shape type is not "GeometricShape".|
||[group](/javascript/api/excel/excel.shapecollectionloadoptions#group)|For EACH ITEM in the collection: Returns the shape group associated with the shape. An error will be thrown if the shape type is not "GroupShape".|
||[height](/javascript/api/excel/excel.shapecollectionloadoptions#height)|For EACH ITEM in the collection: Represents the height, in points, of the shape.|
||[id](/javascript/api/excel/excel.shapecollectionloadoptions#id)|For EACH ITEM in the collection: Represents the shape identifier. Read-only.|
||[image](/javascript/api/excel/excel.shapecollectionloadoptions#image)|For EACH ITEM in the collection: Returns the image associated with the shape. An error will be thrown if the shape type is not "Image".|
||[left](/javascript/api/excel/excel.shapecollectionloadoptions#left)|For EACH ITEM in the collection: The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[level](/javascript/api/excel/excel.shapecollectionloadoptions#level)|For EACH ITEM in the collection: Represents the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.|
||[line](/javascript/api/excel/excel.shapecollectionloadoptions#line)|For EACH ITEM in the collection: Returns the line associated with the shape. An error will be thrown if the shape type is not "Line".|
||[lineFormat](/javascript/api/excel/excel.shapecollectionloadoptions#lineformat)|For EACH ITEM in the collection: Returns the line formatting of this shape.|
||[lockAspectRatio](/javascript/api/excel/excel.shapecollectionloadoptions#lockaspectratio)|For EACH ITEM in the collection: Specifies whether or not the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.shapecollectionloadoptions#name)|For EACH ITEM in the collection: Represents the name of the shape.|
||[parentGroup](/javascript/api/excel/excel.shapecollectionloadoptions#parentgroup)|For EACH ITEM in the collection: Represents the parent group of this shape.|
||[rotation](/javascript/api/excel/excel.shapecollectionloadoptions#rotation)|For EACH ITEM in the collection: Represents the rotation, in degrees, of the shape.|
||[textFrame](/javascript/api/excel/excel.shapecollectionloadoptions#textframe)|For EACH ITEM in the collection: Returns the text frame object of this shape. Read only.|
||[top](/javascript/api/excel/excel.shapecollectionloadoptions#top)|For EACH ITEM in the collection: The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[type](/javascript/api/excel/excel.shapecollectionloadoptions#type)|For EACH ITEM in the collection: Returns the type of this shape. See Excel.ShapeType for details. Read-only.|
||[visible](/javascript/api/excel/excel.shapecollectionloadoptions#visible)|For EACH ITEM in the collection: Represents the visibility of this shape.|
||[width](/javascript/api/excel/excel.shapecollectionloadoptions#width)|For EACH ITEM in the collection: Represents the width, in points, of the shape.|
||[zOrderPosition](/javascript/api/excel/excel.shapecollectionloadoptions#zorderposition)|For EACH ITEM in the collection: Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack. Read-only.|
|[ShapeData](/javascript/api/excel/excel.shapedata)|[altTextDescription](/javascript/api/excel/excel.shapedata#alttextdescription)|Returns or sets the alternative description text for a Shape object.|
||[altTextTitle](/javascript/api/excel/excel.shapedata#alttexttitle)|Returns or sets the alternative title text for a Shape object.|
||[connectionSiteCount](/javascript/api/excel/excel.shapedata#connectionsitecount)|Returns the number of connection sites on this shape. Read-only.|
||[fill](/javascript/api/excel/excel.shapedata#fill)|Returns the fill formatting of this shape. Read-only.|
||[geometricShapeType](/javascript/api/excel/excel.shapedata#geometricshapetype)|Represents the geometric shape type of this geometric shape. See Excel.GeometricShapeType for details. Returns null if the shape type is not "GeometricShape".|
||[height](/javascript/api/excel/excel.shapedata#height)|Represents the height, in points, of the shape.|
||[id](/javascript/api/excel/excel.shapedata#id)|Represents the shape identifier. Read-only.|
||[left](/javascript/api/excel/excel.shapedata#left)|The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[level](/javascript/api/excel/excel.shapedata#level)|Represents the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.|
||[lineFormat](/javascript/api/excel/excel.shapedata#lineformat)|Returns the line formatting of this shape. Read-only.|
||[lockAspectRatio](/javascript/api/excel/excel.shapedata#lockaspectratio)|Specifies whether or not the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.shapedata#name)|Represents the name of the shape.|
||[rotation](/javascript/api/excel/excel.shapedata#rotation)|Represents the rotation, in degrees, of the shape.|
||[top](/javascript/api/excel/excel.shapedata#top)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[type](/javascript/api/excel/excel.shapedata#type)|Returns the type of this shape. See Excel.ShapeType for details. Read-only.|
||[visible](/javascript/api/excel/excel.shapedata#visible)|Represents the visibility of this shape.|
||[width](/javascript/api/excel/excel.shapedata#width)|Represents the width, in points, of the shape.|
||[zOrderPosition](/javascript/api/excel/excel.shapedata#zorderposition)|Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack. Read-only.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Gets the id of the shape deactivated shape.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Gets the id of the worksheet in which the shape is deactivated.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Clears the fill formatting of this shape.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Represents the shape fill foreground color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
||[type](/javascript/api/excel/excel.shapefill#type)|Returns the fill type of the shape. Read-only. See Excel.ShapeFillType for details.|
||[set(properties: Excel.ShapeFill)](/javascript/api/excel/excel.shapefill#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ShapeFillUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.shapefill#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Sets the fill formatting of the shape to a uniform color. This changes the fill type to "Solid".|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Returns or sets the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.|
|[ShapeFillData](/javascript/api/excel/excel.shapefilldata)|[foregroundColor](/javascript/api/excel/excel.shapefilldata#foregroundcolor)|Represents the shape fill foreground color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
||[transparency](/javascript/api/excel/excel.shapefilldata#transparency)|Returns or sets the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.|
||[type](/javascript/api/excel/excel.shapefilldata#type)|Returns the fill type of the shape. Read-only. See Excel.ShapeFillType for details.|
|[ShapeFillLoadOptions](/javascript/api/excel/excel.shapefillloadoptions)|[$all](/javascript/api/excel/excel.shapefillloadoptions#$all)||
||[foregroundColor](/javascript/api/excel/excel.shapefillloadoptions#foregroundcolor)|Represents the shape fill foreground color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
||[transparency](/javascript/api/excel/excel.shapefillloadoptions#transparency)|Returns or sets the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.|
||[type](/javascript/api/excel/excel.shapefillloadoptions#type)|Returns the fill type of the shape. Read-only. See Excel.ShapeFillType for details.|
|[ShapeFillUpdateData](/javascript/api/excel/excel.shapefillupdatedata)|[foregroundColor](/javascript/api/excel/excel.shapefillupdatedata#foregroundcolor)|Represents the shape fill foreground color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
||[transparency](/javascript/api/excel/excel.shapefillupdatedata#transparency)|Returns or sets the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|
||[color](/javascript/api/excel/excel.shapefont#color)|The HTML color code representation of the text color (e.g. "#FF0000" represents red). Returns null if the TextRange includes text fragments with different colors.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Represents the italic status of font. Returns null if the TextRange includes both italic and non-italic text fragments.|
||[name](/javascript/api/excel/excel.shapefont#name)|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.|
||[set(properties: Excel.ShapeFont)](/javascript/api/excel/excel.shapefont#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ShapeFontUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.shapefont#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[size](/javascript/api/excel/excel.shapefont#size)|Represents font size in points (e.g. 11). Returns null if the TextRange includes text fragments with different font sizes.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Type of underline applied to the font. Returns null if the TextRange includes text fragments with different underline styles. See Excel.ShapeFontUnderlineStyle for details.|
|[ShapeFontData](/javascript/api/excel/excel.shapefontdata)|[bold](/javascript/api/excel/excel.shapefontdata#bold)|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|
||[color](/javascript/api/excel/excel.shapefontdata#color)|The HTML color code representation of the text color (e.g. "#FF0000" represents red). Returns null if the TextRange includes text fragments with different colors.|
||[italic](/javascript/api/excel/excel.shapefontdata#italic)|Represents the italic status of font. Returns null if the TextRange includes both italic and non-italic text fragments.|
||[name](/javascript/api/excel/excel.shapefontdata#name)|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.|
||[size](/javascript/api/excel/excel.shapefontdata#size)|Represents font size in points (e.g. 11). Returns null if the TextRange includes text fragments with different font sizes.|
||[underline](/javascript/api/excel/excel.shapefontdata#underline)|Type of underline applied to the font. Returns null if the TextRange includes text fragments with different underline styles. See Excel.ShapeFontUnderlineStyle for details.|
|[ShapeFontLoadOptions](/javascript/api/excel/excel.shapefontloadoptions)|[$all](/javascript/api/excel/excel.shapefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.shapefontloadoptions#bold)|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|
||[color](/javascript/api/excel/excel.shapefontloadoptions#color)|The HTML color code representation of the text color (e.g. "#FF0000" represents red). Returns null if the TextRange includes text fragments with different colors.|
||[italic](/javascript/api/excel/excel.shapefontloadoptions#italic)|Represents the italic status of font. Returns null if the TextRange includes both italic and non-italic text fragments.|
||[name](/javascript/api/excel/excel.shapefontloadoptions#name)|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.|
||[size](/javascript/api/excel/excel.shapefontloadoptions#size)|Represents font size in points (e.g. 11). Returns null if the TextRange includes text fragments with different font sizes.|
||[underline](/javascript/api/excel/excel.shapefontloadoptions#underline)|Type of underline applied to the font. Returns null if the TextRange includes text fragments with different underline styles. See Excel.ShapeFontUnderlineStyle for details.|
|[ShapeFontUpdateData](/javascript/api/excel/excel.shapefontupdatedata)|[bold](/javascript/api/excel/excel.shapefontupdatedata#bold)|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|
||[color](/javascript/api/excel/excel.shapefontupdatedata#color)|The HTML color code representation of the text color (e.g. "#FF0000" represents red). Returns null if the TextRange includes text fragments with different colors.|
||[italic](/javascript/api/excel/excel.shapefontupdatedata#italic)|Represents the italic status of font. Returns null if the TextRange includes both italic and non-italic text fragments.|
||[name](/javascript/api/excel/excel.shapefontupdatedata#name)|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.|
||[size](/javascript/api/excel/excel.shapefontupdatedata#size)|Represents font size in points (e.g. 11). Returns null if the TextRange includes text fragments with different font sizes.|
||[underline](/javascript/api/excel/excel.shapefontupdatedata#underline)|Type of underline applied to the font. Returns null if the TextRange includes text fragments with different underline styles. See Excel.ShapeFontUnderlineStyle for details.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Represents the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Returns the Shape object associated with the group. Read-only.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Returns the collection of Shape objects. Read-only.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Ungroups any grouped shapes in the specified shape group.|
|[ShapeGroupData](/javascript/api/excel/excel.shapegroupdata)|[id](/javascript/api/excel/excel.shapegroupdata#id)|Represents the shape identifier. Read-only.|
||[shapes](/javascript/api/excel/excel.shapegroupdata#shapes)|Returns the collection of Shape objects. Read-only.|
|[ShapeGroupLoadOptions](/javascript/api/excel/excel.shapegrouploadoptions)|[$all](/javascript/api/excel/excel.shapegrouploadoptions#$all)||
||[id](/javascript/api/excel/excel.shapegrouploadoptions#id)|Represents the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.shapegrouploadoptions#shape)|Returns the Shape object associated with the group.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent dash styles. See Excel.ShapeLineStyle for details.|
||[set(properties: Excel.ShapeLineFormat)](/javascript/api/excel/excel.shapelineformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ShapeLineFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.shapelineformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See Excel.ShapeLineStyle for details.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Represents whether or not the line formatting of a shape element is visible. Returns null when the shape has inconsistent visibilities.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Represents the weight of the line, in points. Returns null when the line is not visible or there are inconsistent line weights.|
|[ShapeLineFormatData](/javascript/api/excel/excel.shapelineformatdata)|[color](/javascript/api/excel/excel.shapelineformatdata#color)|Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[dashStyle](/javascript/api/excel/excel.shapelineformatdata#dashstyle)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent dash styles. See Excel.ShapeLineStyle for details.|
||[style](/javascript/api/excel/excel.shapelineformatdata#style)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See Excel.ShapeLineStyle for details.|
||[transparency](/javascript/api/excel/excel.shapelineformatdata#transparency)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.|
||[visible](/javascript/api/excel/excel.shapelineformatdata#visible)|Represents whether or not the line formatting of a shape element is visible. Returns null when the shape has inconsistent visibilities.|
||[weight](/javascript/api/excel/excel.shapelineformatdata#weight)|Represents the weight of the line, in points. Returns null when the line is not visible or there are inconsistent line weights.|
|[ShapeLineFormatLoadOptions](/javascript/api/excel/excel.shapelineformatloadoptions)|[$all](/javascript/api/excel/excel.shapelineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.shapelineformatloadoptions#color)|Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[dashStyle](/javascript/api/excel/excel.shapelineformatloadoptions#dashstyle)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent dash styles. See Excel.ShapeLineStyle for details.|
||[style](/javascript/api/excel/excel.shapelineformatloadoptions#style)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See Excel.ShapeLineStyle for details.|
||[transparency](/javascript/api/excel/excel.shapelineformatloadoptions#transparency)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.|
||[visible](/javascript/api/excel/excel.shapelineformatloadoptions#visible)|Represents whether or not the line formatting of a shape element is visible. Returns null when the shape has inconsistent visibilities.|
||[weight](/javascript/api/excel/excel.shapelineformatloadoptions#weight)|Represents the weight of the line, in points. Returns null when the line is not visible or there are inconsistent line weights.|
|[ShapeLineFormatUpdateData](/javascript/api/excel/excel.shapelineformatupdatedata)|[color](/javascript/api/excel/excel.shapelineformatupdatedata#color)|Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[dashStyle](/javascript/api/excel/excel.shapelineformatupdatedata#dashstyle)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent dash styles. See Excel.ShapeLineStyle for details.|
||[style](/javascript/api/excel/excel.shapelineformatupdatedata#style)|Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See Excel.ShapeLineStyle for details.|
||[transparency](/javascript/api/excel/excel.shapelineformatupdatedata#transparency)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.|
||[visible](/javascript/api/excel/excel.shapelineformatupdatedata#visible)|Represents whether or not the line formatting of a shape element is visible. Returns null when the shape has inconsistent visibilities.|
||[weight](/javascript/api/excel/excel.shapelineformatupdatedata#weight)|Represents the weight of the line, in points. Returns null when the line is not visible or there are inconsistent line weights.|
|[ShapeLoadOptions](/javascript/api/excel/excel.shapeloadoptions)|[$all](/javascript/api/excel/excel.shapeloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapeloadoptions#alttextdescription)|Returns or sets the alternative description text for a Shape object.|
||[altTextTitle](/javascript/api/excel/excel.shapeloadoptions#alttexttitle)|Returns or sets the alternative title text for a Shape object.|
||[connectionSiteCount](/javascript/api/excel/excel.shapeloadoptions#connectionsitecount)|Returns the number of connection sites on this shape. Read-only.|
||[fill](/javascript/api/excel/excel.shapeloadoptions#fill)|Returns the fill formatting of this shape.|
||[geometricShape](/javascript/api/excel/excel.shapeloadoptions#geometricshape)|Returns the geometric shape associated with the shape. An error will be thrown if the shape type is not "GeometricShape".|
||[geometricShapeType](/javascript/api/excel/excel.shapeloadoptions#geometricshapetype)|Represents the geometric shape type of this geometric shape. See Excel.GeometricShapeType for details. Returns null if the shape type is not "GeometricShape".|
||[group](/javascript/api/excel/excel.shapeloadoptions#group)|Returns the shape group associated with the shape. An error will be thrown if the shape type is not "GroupShape".|
||[height](/javascript/api/excel/excel.shapeloadoptions#height)|Represents the height, in points, of the shape.|
||[id](/javascript/api/excel/excel.shapeloadoptions#id)|Represents the shape identifier. Read-only.|
||[image](/javascript/api/excel/excel.shapeloadoptions#image)|Returns the image associated with the shape. An error will be thrown if the shape type is not "Image".|
||[left](/javascript/api/excel/excel.shapeloadoptions#left)|The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[level](/javascript/api/excel/excel.shapeloadoptions#level)|Represents the level of the specified shape. For example, a level of 0 means that the shape is not part of any groups, a level of 1 means the shape is part of a top-level group, and a level of 2 means the shape is part of a sub-group of the top level.|
||[line](/javascript/api/excel/excel.shapeloadoptions#line)|Returns the line associated with the shape. An error will be thrown if the shape type is not "Line".|
||[lineFormat](/javascript/api/excel/excel.shapeloadoptions#lineformat)|Returns the line formatting of this shape.|
||[lockAspectRatio](/javascript/api/excel/excel.shapeloadoptions#lockaspectratio)|Specifies whether or not the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.shapeloadoptions#name)|Represents the name of the shape.|
||[parentGroup](/javascript/api/excel/excel.shapeloadoptions#parentgroup)|Represents the parent group of this shape.|
||[rotation](/javascript/api/excel/excel.shapeloadoptions#rotation)|Represents the rotation, in degrees, of the shape.|
||[textFrame](/javascript/api/excel/excel.shapeloadoptions#textframe)|Returns the text frame object of this shape. Read only.|
||[top](/javascript/api/excel/excel.shapeloadoptions#top)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[type](/javascript/api/excel/excel.shapeloadoptions#type)|Returns the type of this shape. See Excel.ShapeType for details. Read-only.|
||[visible](/javascript/api/excel/excel.shapeloadoptions#visible)|Represents the visibility of this shape.|
||[width](/javascript/api/excel/excel.shapeloadoptions#width)|Represents the width, in points, of the shape.|
||[zOrderPosition](/javascript/api/excel/excel.shapeloadoptions#zorderposition)|Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack. Read-only.|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[altTextDescription](/javascript/api/excel/excel.shapeupdatedata#alttextdescription)|Returns or sets the alternative description text for a Shape object.|
||[altTextTitle](/javascript/api/excel/excel.shapeupdatedata#alttexttitle)|Returns or sets the alternative title text for a Shape object.|
||[fill](/javascript/api/excel/excel.shapeupdatedata#fill)|Returns the fill formatting of this shape.|
||[geometricShapeType](/javascript/api/excel/excel.shapeupdatedata#geometricshapetype)|Represents the geometric shape type of this geometric shape. See Excel.GeometricShapeType for details. Returns null if the shape type is not "GeometricShape".|
||[height](/javascript/api/excel/excel.shapeupdatedata#height)|Represents the height, in points, of the shape.|
||[left](/javascript/api/excel/excel.shapeupdatedata#left)|The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[lineFormat](/javascript/api/excel/excel.shapeupdatedata#lineformat)|Returns the line formatting of this shape.|
||[lockAspectRatio](/javascript/api/excel/excel.shapeupdatedata#lockaspectratio)|Specifies whether or not the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.shapeupdatedata#name)|Represents the name of the shape.|
||[rotation](/javascript/api/excel/excel.shapeupdatedata#rotation)|Represents the rotation, in degrees, of the shape.|
||[top](/javascript/api/excel/excel.shapeupdatedata#top)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[visible](/javascript/api/excel/excel.shapeupdatedata#visible)|Represents the visibility of this shape.|
||[width](/javascript/api/excel/excel.shapeupdatedata#width)|Represents the width, in points, of the shape.|
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
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.tablecollectionloadoptions#autofilter)|For EACH ITEM in the collection: Represents the AutoFilter object of the table.|
|[TableData](/javascript/api/excel/excel.tabledata)|[autoFilter](/javascript/api/excel/excel.tabledata#autofilter)|Represents the AutoFilter object of the table. Read-Only.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Specifies the source of the event. See Excel.EventSource for details.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Specifies the id of the table that is deleted.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Specifies the name of the table that is deleted.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Specifies the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Specifies the id of the worksheet in which the table is deleted.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[autoFilter](/javascript/api/excel/excel.tableloadoptions#autofilter)|Represents the AutoFilter object of the table.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Gets the number of tables in the collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Gets the first table in the collection. The tables in the collection are sorted top to bottom and left to right, such that top left table is the first table in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Gets a table by Name or ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Gets the loaded child items in this collection.|
|[TableScopedCollectionLoadOptions](/javascript/api/excel/excel.tablescopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablescopedcollectionloadoptions#$all)||
||[autoFilter](/javascript/api/excel/excel.tablescopedcollectionloadoptions#autofilter)|For EACH ITEM in the collection: Represents the AutoFilter object of the table.|
||[columns](/javascript/api/excel/excel.tablescopedcollectionloadoptions#columns)|For EACH ITEM in the collection: Represents a collection of all the columns in the table.|
||[highlightFirstColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightfirstcolumn)|For EACH ITEM in the collection: Indicates whether the first column contains special formatting.|
||[highlightLastColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightlastcolumn)|For EACH ITEM in the collection: Indicates whether the last column contains special formatting.|
||[id](/javascript/api/excel/excel.tablescopedcollectionloadoptions#id)|For EACH ITEM in the collection: Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
||[legacyId](/javascript/api/excel/excel.tablescopedcollectionloadoptions#legacyid)|For EACH ITEM in the collection: Returns a numeric id.|
||[name](/javascript/api/excel/excel.tablescopedcollectionloadoptions#name)|For EACH ITEM in the collection: Name of the table.|
||[rows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#rows)|For EACH ITEM in the collection: Represents a collection of all the rows in the table.|
||[showBandedColumns](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedcolumns)|For EACH ITEM in the collection: Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedrows)|For EACH ITEM in the collection: Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showfilterbutton)|For EACH ITEM in the collection: Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|
||[showHeaders](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showheaders)|For EACH ITEM in the collection: Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
||[showTotals](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showtotals)|For EACH ITEM in the collection: Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
||[sort](/javascript/api/excel/excel.tablescopedcollectionloadoptions#sort)|For EACH ITEM in the collection: Represents the sorting for the table.|
||[style](/javascript/api/excel/excel.tablescopedcollectionloadoptions#style)|For EACH ITEM in the collection: Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|
||[worksheet](/javascript/api/excel/excel.tablescopedcollectionloadoptions#worksheet)|For EACH ITEM in the collection: The worksheet containing the current table.|
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
||[set(properties: Excel.TextFrame)](/javascript/api/excel/excel.textframe#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TextFrameUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.textframe#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Represents the vertical alignment of the text frame. See Excel.ShapeTextVerticalAlignment for details.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Represents the vertical overflow behavior of the text frame. See Excel.ShapeTextVerticalOverflow for details.|
|[TextFrameData](/javascript/api/excel/excel.textframedata)|[autoSizeSetting](/javascript/api/excel/excel.textframedata#autosizesetting)|Gets or sets the automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.|
||[bottomMargin](/javascript/api/excel/excel.textframedata#bottommargin)|Represents the bottom margin, in points, of the text frame.|
||[hasText](/javascript/api/excel/excel.textframedata#hastext)|Specifies whether the text frame contains text.|
||[horizontalAlignment](/javascript/api/excel/excel.textframedata#horizontalalignment)|Represents the horizontal alignment of the text frame. See Excel.ShapeTextHorizontalAlignment for details.|
||[horizontalOverflow](/javascript/api/excel/excel.textframedata#horizontaloverflow)|Represents the horizontal overflow behavior of the text frame. See Excel.ShapeTextHorizontalOverflow for details.|
||[leftMargin](/javascript/api/excel/excel.textframedata#leftmargin)|Represents the left margin, in points, of the text frame.|
||[orientation](/javascript/api/excel/excel.textframedata#orientation)|Represents the text orientation of the text frame. See Excel.ShapeTextOrientation for details.|
||[readingOrder](/javascript/api/excel/excel.textframedata#readingorder)|Represents the reading order of the text frame, either left-to-right or right-to-left. See Excel.ShapeTextReadingOrder for details.|
||[rightMargin](/javascript/api/excel/excel.textframedata#rightmargin)|Represents the right margin, in points, of the text frame.|
||[topMargin](/javascript/api/excel/excel.textframedata#topmargin)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/excel/excel.textframedata#verticalalignment)|Represents the vertical alignment of the text frame. See Excel.ShapeTextVerticalAlignment for details.|
||[verticalOverflow](/javascript/api/excel/excel.textframedata#verticaloverflow)|Represents the vertical overflow behavior of the text frame. See Excel.ShapeTextVerticalOverflow for details.|
|[TextFrameLoadOptions](/javascript/api/excel/excel.textframeloadoptions)|[$all](/javascript/api/excel/excel.textframeloadoptions#$all)||
||[autoSizeSetting](/javascript/api/excel/excel.textframeloadoptions#autosizesetting)|Gets or sets the automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.|
||[bottomMargin](/javascript/api/excel/excel.textframeloadoptions#bottommargin)|Represents the bottom margin, in points, of the text frame.|
||[hasText](/javascript/api/excel/excel.textframeloadoptions#hastext)|Specifies whether the text frame contains text.|
||[horizontalAlignment](/javascript/api/excel/excel.textframeloadoptions#horizontalalignment)|Represents the horizontal alignment of the text frame. See Excel.ShapeTextHorizontalAlignment for details.|
||[horizontalOverflow](/javascript/api/excel/excel.textframeloadoptions#horizontaloverflow)|Represents the horizontal overflow behavior of the text frame. See Excel.ShapeTextHorizontalOverflow for details.|
||[leftMargin](/javascript/api/excel/excel.textframeloadoptions#leftmargin)|Represents the left margin, in points, of the text frame.|
||[orientation](/javascript/api/excel/excel.textframeloadoptions#orientation)|Represents the text orientation of the text frame. See Excel.ShapeTextOrientation for details.|
||[readingOrder](/javascript/api/excel/excel.textframeloadoptions#readingorder)|Represents the reading order of the text frame, either left-to-right or right-to-left. See Excel.ShapeTextReadingOrder for details.|
||[rightMargin](/javascript/api/excel/excel.textframeloadoptions#rightmargin)|Represents the right margin, in points, of the text frame.|
||[textRange](/javascript/api/excel/excel.textframeloadoptions#textrange)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See Excel.TextRange for details.|
||[topMargin](/javascript/api/excel/excel.textframeloadoptions#topmargin)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/excel/excel.textframeloadoptions#verticalalignment)|Represents the vertical alignment of the text frame. See Excel.ShapeTextVerticalAlignment for details.|
||[verticalOverflow](/javascript/api/excel/excel.textframeloadoptions#verticaloverflow)|Represents the vertical overflow behavior of the text frame. See Excel.ShapeTextVerticalOverflow for details.|
|[TextFrameUpdateData](/javascript/api/excel/excel.textframeupdatedata)|[autoSizeSetting](/javascript/api/excel/excel.textframeupdatedata#autosizesetting)|Gets or sets the automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.|
||[bottomMargin](/javascript/api/excel/excel.textframeupdatedata#bottommargin)|Represents the bottom margin, in points, of the text frame.|
||[horizontalAlignment](/javascript/api/excel/excel.textframeupdatedata#horizontalalignment)|Represents the horizontal alignment of the text frame. See Excel.ShapeTextHorizontalAlignment for details.|
||[horizontalOverflow](/javascript/api/excel/excel.textframeupdatedata#horizontaloverflow)|Represents the horizontal overflow behavior of the text frame. See Excel.ShapeTextHorizontalOverflow for details.|
||[leftMargin](/javascript/api/excel/excel.textframeupdatedata#leftmargin)|Represents the left margin, in points, of the text frame.|
||[orientation](/javascript/api/excel/excel.textframeupdatedata#orientation)|Represents the text orientation of the text frame. See Excel.ShapeTextOrientation for details.|
||[readingOrder](/javascript/api/excel/excel.textframeupdatedata#readingorder)|Represents the reading order of the text frame, either left-to-right or right-to-left. See Excel.ShapeTextReadingOrder for details.|
||[rightMargin](/javascript/api/excel/excel.textframeupdatedata#rightmargin)|Represents the right margin, in points, of the text frame.|
||[topMargin](/javascript/api/excel/excel.textframeupdatedata#topmargin)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/excel/excel.textframeupdatedata#verticalalignment)|Represents the vertical alignment of the text frame. See Excel.ShapeTextVerticalAlignment for details.|
||[verticalOverflow](/javascript/api/excel/excel.textframeupdatedata#verticaloverflow)|Represents the vertical overflow behavior of the text frame. See Excel.ShapeTextVerticalOverflow for details.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Returns a TextRange object for the substring in the given range.|
||[font](/javascript/api/excel/excel.textrange#font)|Returns a ShapeFont object that represents the font attributes for the text range. Read-only.|
||[set(properties: Excel.TextRange)](/javascript/api/excel/excel.textrange#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TextRangeUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.textrange#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[text](/javascript/api/excel/excel.textrange#text)|Represents the plain text content of the text range.|
|[TextRangeData](/javascript/api/excel/excel.textrangedata)|[font](/javascript/api/excel/excel.textrangedata#font)|Returns a ShapeFont object that represents the font attributes for the text range. Read-only.|
||[text](/javascript/api/excel/excel.textrangedata#text)|Represents the plain text content of the text range.|
|[TextRangeLoadOptions](/javascript/api/excel/excel.textrangeloadoptions)|[$all](/javascript/api/excel/excel.textrangeloadoptions#$all)||
||[font](/javascript/api/excel/excel.textrangeloadoptions#font)|Returns a ShapeFont object that represents the font attributes for the text range.|
||[text](/javascript/api/excel/excel.textrangeloadoptions#text)|Represents the plain text content of the text range.|
|[TextRangeUpdateData](/javascript/api/excel/excel.textrangeupdatedata)|[font](/javascript/api/excel/excel.textrangeupdatedata#font)|Returns a ShapeFont object that represents the font attributes for the text range.|
||[text](/javascript/api/excel/excel.textrangeupdatedata#text)|Represents the plain text content of the text range.|
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
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[autoSave](/javascript/api/excel/excel.workbookdata#autosave)|Specifies whether or not the workbook is in autosave mode. Read-Only.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookdata#calculationengineversion)|Returns a number about the version of Excel Calculation Engine. Read-Only.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookdata#chartdatapointtrack)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[isDirty](/javascript/api/excel/excel.workbookdata#isdirty)|Specifies whether or not changes have been made since the workbook was last saved.|
||[previouslySaved](/javascript/api/excel/excel.workbookdata#previouslysaved)|Specifies whether or not the workbook has ever been saved locally or online. Read-Only.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookdata#useprecisionasdisplayed)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[autoSave](/javascript/api/excel/excel.workbookloadoptions#autosave)|Specifies whether or not the workbook is in autosave mode. Read-Only.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookloadoptions#calculationengineversion)|Returns a number about the version of Excel Calculation Engine. Read-Only.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookloadoptions#chartdatapointtrack)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[isDirty](/javascript/api/excel/excel.workbookloadoptions#isdirty)|Specifies whether or not changes have been made since the workbook was last saved.|
||[previouslySaved](/javascript/api/excel/excel.workbookloadoptions#previouslysaved)|Specifies whether or not the workbook has ever been saved locally or online. Read-Only.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookloadoptions#useprecisionasdisplayed)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[chartDataPointTrack](/javascript/api/excel/excel.workbookupdatedata#chartdatapointtrack)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[isDirty](/javascript/api/excel/excel.workbookupdatedata#isdirty)|Specifies whether or not changes have been made since the workbook was last saved.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookupdatedata#useprecisionasdisplayed)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
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
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetcollectionloadoptions#autofilter)|For EACH ITEM in the collection: Represents the AutoFilter object of the worksheet.|
||[enableCalculation](/javascript/api/excel/excel.worksheetcollectionloadoptions#enablecalculation)|For EACH ITEM in the collection: Gets or sets the enableCalculation property of the worksheet.|
||[pageLayout](/javascript/api/excel/excel.worksheetcollectionloadoptions#pagelayout)|For EACH ITEM in the collection: Gets the PageLayout object of the worksheet.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[autoFilter](/javascript/api/excel/excel.worksheetdata#autofilter)|Represents the AutoFilter object of the worksheet. Read-Only.|
||[enableCalculation](/javascript/api/excel/excel.worksheetdata#enablecalculation)|Gets or sets the enableCalculation property of the worksheet.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheetdata#horizontalpagebreaks)|Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks.|
||[pageLayout](/javascript/api/excel/excel.worksheetdata#pagelayout)|Gets the PageLayout object of the worksheet.|
||[shapes](/javascript/api/excel/excel.worksheetdata#shapes)|Returns the collection of all the Shape objects on the worksheet. Read-only.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheetdata#verticalpagebreaks)|Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Gets the range that represents the changed area of a specific worksheet. It might return null object.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetloadoptions#autofilter)|Represents the AutoFilter object of the worksheet.|
||[enableCalculation](/javascript/api/excel/excel.worksheetloadoptions#enablecalculation)|Gets or sets the enableCalculation property of the worksheet.|
||[pageLayout](/javascript/api/excel/excel.worksheetloadoptions#pagelayout)|Gets the PageLayout object of the worksheet.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Specifies whether the match needs to be complete or partial. A complete match matches the entire contents of the cell. Default is false (partial).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Specifies whether the match is case sensitive. Default is false (insensitive).|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[enableCalculation](/javascript/api/excel/excel.worksheetupdatedata#enablecalculation)|Gets or sets the enableCalculation property of the worksheet.|
||[pageLayout](/javascript/api/excel/excel.worksheetupdatedata#pagelayout)|Gets the PageLayout object of the worksheet.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
