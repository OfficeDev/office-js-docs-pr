---
title: Excel JavaScript API requirement sets
description: ''
ms.date: 10/09/2018
ms.prod: excel
localization_priority: Priority
---

# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Excel add-ins run across multiple versions of Office, including Office 2016 or later for Windows, Office for iPad, Office for Mac, and Office Online. The following table lists the Excel requirement sets, the Office host applications that support each requirement set, and the build versions or number for those applications.

> [!NOTE]
> To use APIs in any of the numbered requirement sets, you should reference the **production** library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> For information about using preview APIs, see the [Excel JavaScript preview APIs](#excel-javascript-preview-apis) section within this article.

|  Requirement set  |  Office 365 for Windows  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Preview  | Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://products.office.com/office-insider)) |
| ExcelApi1.8  | Version 1808 (Build 10730.20102) or later | 2.17 or later | 16.17 or later | September 2018 | Coming soon |
| ExcelApi1.7  | Version 1801 (Build 9001.2171) or later   | 2.9 or later | 16.9 or later | April 2018 | Coming soon |
| ExcelApi1.6  | Version 1704 (Build 8201.2001) or later   | 2.2 or later |15.36 or later| April 2017 | Coming soon|
| ExcelApi1.5  | Version 1703 (Build 8067.2070) or later   | 2.2 or later |15.36 or later| March 2017 | Coming soon|
| ExcelApi1.4  | Version 1701 (Build 7870.2024) or later   | 2.2 or later |15.36 or later| January 2017 | Coming soon|
| ExcelApi1.3  | Version 1608 (Build 7369.2055) or later | 1.27 or later |  15.27 or later| September 2016 | Version 1608 (Build 7601.6800) or later|
| ExcelApi1.2  | Version 1601 (Build 6741.2088) or later | 1.21 or later | 15.22 or later| January 2016 ||
| ExcelApi1.1  | Version 1509 (Build 4266.1001) or later | 1.19 or later | 15.20 or later| January 2016 ||

> [!NOTE]
> The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1 requirement set.

For more information about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired. The following table lists the APIs currently available in preview. To provide feedback about a preview API, please use the feedback mechanism at the end of the web page where the API is documented.

> [!NOTE]
> Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.
>
> To use preview APIs, you must reference the **beta** library on the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js and you may also need to join the Office Insider program to get a sufficiently recent Office build.

More than 400 new Excel APIs are currently in preview. The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list. Please try the new features and share your feedback with us.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Slicer | Insert and configure slicers to tables and PivotTables. | [Slicer](/javascript/api/excel/slicer) |
| Comments | Add, edit, and delete comments. | [Comment](/javascript/api/excel/comment), [CommentCollection](/javascript/api/excel/commentcollection) |
| Shapes | Insert, position, and format images, geometric shapes and text boxes. | [ShapeCollection](/javascript/api/excel/shapecollection) [Shape](/javascript/api/excel/shape) [GeometricShape](/javascript/api/excel/geometricshape)  [Image](/javascript/api/excel/image) |
| New Charts | Explore our new supported chart types: maps, box and whisker, waterfall, sunburst, pareto. and funnel. | [Chart](/javascript/api/excel/charttype) |
| Auto Filter | Add filters to ranges. | [AutoFilter](/javascript/api/excel/autofilter) |
| Areas | Support for discontinuous ranges. | [RangeAreas](/javascript/api/excel/rangeareas) |
| Special Cells | Get cells containing dates, comments, or formulas within a range. | [Range](/javascript/api/excel/range#getspecialcells-celltype--cellvaluetype-)|
| Find | Find values or formulas within a range or worksheet. | [Range](/javascript/api/excel/range#find-text--criteria-)[Worksheet](/javascript/api/excel/worksheet#findall-text--criteria-) |
| Copy Paste | Copy values, formats, and formulas from one range to another. | [Range](/javascript/api/excel/range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| RangeFormat | New capabilities with range formats. | [Range](/javascript/api/excel/rangeformat) |
| Workbook Save, Close | Save and close workbooks.  | [Workbook](/javascript/api/excel/workbook) |
| Insert Workbook | Insert one workbook into another.  | [Workbook](/javascript/api/excel/worksheetcollection) |
| Calculation | Greater control over the Excel calculation engine. | [Application](/javascript/api/excel/application) |

The following is a complete list of APIs in preview.

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Returns a number about the version of Excel Calculation Engine that the workbook was last fully recalculated by. Read-only.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Returns a CalculationState that indicates the calculation state of the application. See Excel.CalculationState for details. Read-only.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Returns the Iterative Calculation settings.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Suspends sceen updating until the next "context.sync()" is called.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Applies AutoFilter on a range and filters the column if column index and filter criteria are specified.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Clears the criteria if AutoFilter has filters|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Returns the Range object that represents the range to which the AutoFilter applies.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|If there is Range object associated with the AutoFilter, this method returns it.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Array that holds all filter criteria in an autofiltered range. Read-Only.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Indicates if the AutoFilter is enabled or not. Read-Only.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Indicates if the AutoFilter has filter criteria. Read-Only.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Applies the specified Autofilter object currently on the range.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Removes the AutoFilter for the range.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)||
||[style](/javascript/api/excel/excel.cellborder#style)||
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)||
||[weight](/javascript/api/excel/excel.cellborder#weight)||
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)||
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)||
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)||
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)||
||[left](/javascript/api/excel/excel.cellbordercollection#left)||
||[right](/javascript/api/excel/excel.cellbordercollection#right)||
||[top](/javascript/api/excel/excel.cellbordercollection#top)||
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)||
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)||
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)||
||[hasSpill](/javascript/api/excel/excel.cellproperties#hasspill)||
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)||
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)||
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)||
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)||
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)||
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)||
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)||
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)||
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)||
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)||
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)||
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)||
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)||
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)||
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)||
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)||
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)||
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)||
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)||
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)||
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)||
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)||
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)||
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)||
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)||
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)||
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)||
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)||
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)||
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Creates and opens a new workbook.  Optionally, the workbook can be pre-populated with a base64-encoded .xlsx file.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)||
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)||
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Activate the chart in the Excel UI.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Encapsulates the options for the pivot chart. Read-only.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Returns or sets an integer that represents the color scheme for the chart. Read/Write.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|True if the chart area of the chart has rounded corners. Read/Write.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Returns or sets if bin overflow enabled in a histogram chart or pareto chart. Read/Write.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Returns or sets if bin underflow enabled in a histogram chart or pareto chart. Read/Write.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Returns or sets count of bin of a histogram chart or pareto chart. Read/Write.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Returns or sets bin overflow value of a histogram chart or pareto chart. Read/Write.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Returns or sets bin type of a histogram chart or pareto chart. Read/Write.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Returns or sets bin underflow value of a histogram chart or pareto chart. Read/Write.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Returns or sets bin width value of a histogram chart or pareto chart. Read/Write.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Returns or sets quartile calculation type of a Box & whisker chart. Read/Write.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Returns or sets if inner points showed in a Box & whisker chart. Read/Write.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Returns or sets if mean line showed in a Box & whisker chart. Read/Write.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Returns or sets if mean marker showed in a Box & whisker chart. Read/Write.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Returns or sets if outlier points showed in a Box & whisker chart. Read/Write.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Represents whether have the end style cap for the error bars.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Represents which error-bar parts to include. See Excel.ChartErrorBarsInclude for details.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Represents the formatting of chart ErrorBars.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|Represents the range marked by error bars. See Excel.ChartErrorBarsType for details.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Represents whether shown error bars.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Represents chart line formatting.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Returns or sets series map labels strategy of a region map chart. Read/Write.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Returns or sets series map area of a region map chart. Read/Write.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Returns or sets series projection type of a region map chart. Read/Write.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Represents whether to display axis field buttons on a PivotChart.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Represents whether to display legend field buttons on a PivotChart.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Represents whether to display report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Represents whether to display show value field buttons on a PivotChart.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. Read/Write.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Returns or sets the Color for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Returns or sets the type for maximum value of a region map chart series. Read/Write.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Returns or sets the maximum value of a region map chart series. Read/Write.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Returns or sets the Color for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Returns or sets the type for midpoint value of a region map chart series. Read/Write.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Returns or sets the midpoint value of a region map chart series. Read/Write.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Returns or sets the Color for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Returns or sets the type for minimum value of a region map chart series. Read/Write.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Returns or sets the minimum value of a region map chart series. Read/Write.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Returns or sets series gradient style of a region map chart. Read/Write.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Returns or sets the fill color for negative data points in a series. Read/Write.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Returns or sets series parent label strategy area of a treemap chart. Read/Write.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Encapsulates the bin options only for histogram chart and pareto chart. Read-only.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Encapsulates the options for the Box & Whisker chart. Read-only.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Encapsulates the options for the Map chart. Read-only.|
||[xerrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Represents the error bar object for a chart series.|
||[yerrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Represents the error bar object for a chart series.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Returns or sets if connector lines show in a waterfall chart. Read/Write.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|True if Microsoft Excel show leaderlines for each datalabel in series. Read/Write.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. Read/Write.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)||
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)||
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)||
||[hasSpill](/javascript/api/excel/excel.columnproperties#hasspill)||
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Get/Set the content.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Deletes the comment thread.|
||[id](/javascript/api/excel/excel.comment#id)|Represents the comment identifier. Read-only.|
||[isParent](/javascript/api/excel/excel.comment#isparent)|Represents whether it is a comment thread or reply. Always return true here. Read-only.|
||[replies](/javascript/api/excel/excel.comment#replies)|Represents a collection of reply objects associated with the comment. Read-only.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Creates a new comment(comment thread) based on the cell location and content. Invalid argument will be thrown if the location is larger than one cell.|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Creates a new comment(comment thread) based on the cell location and content. Invalid argument will be thrown if the location is larger than one cell.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Gets the number of comments in the collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Returns a comment identified by its ID. Read-only.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Gets a comment based on its position in the collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Gets a comment on the specific cell in the collection.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Gets a comment related to its reply ID in the collection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.commentcollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Gets the loaded child items in this collection.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Get/Set the content.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Deletes the comment reply.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Get its parent comment of this reply.|
||[id](/javascript/api/excel/excel.commentreply#id)|Represents the comment reply identifier. Read-only.|
||[isParent](/javascript/api/excel/excel.commentreply#isparent)|Represents whether it is a comment thread or reply. Always return false here. Read-only.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Creates a comment reply for comment.|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Creates a comment reply for comment.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Gets the number of comment replies in the collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Returns a comment reply identified by its ID. Read-only.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Gets a comment reply based on its position in the collection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.commentreplycollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Gets the loaded child items in this collection.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Returns the RangeAreas, comprising one or more rectangular ranges, the conditonal format is applied to. Read-only.|
|[CustomFunctionEventArgs](/javascript/api/excel/excel.customfunctioneventargs)|[higherTicks](/javascript/api/excel/excel.customfunctioneventargs#higherticks)||
||[lowerTicks](/javascript/api/excel/excel.customfunctioneventargs#lowerticks)||
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will throw an ItemNotFound error.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will return null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|The property used by the filter to do rich filter on richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Represents the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Returns the shape object for the geometric shape. Read-only.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Returns the number of shapes in the group shape. Read-only.|
||[getItem(name: string)](/javascript/api/excel/excel.groupshapecollection#getitem-name-)|Gets a shape using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Gets a shape based on its position in the collection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.groupshapecollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
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
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Returns the format for the image. Read-only.|
||[id](/javascript/api/excel/excel.image#id)|Represents the shape identifier for the image object. Read-only.|
||[shape](/javascript/api/excel/excel.image#shape)|Returns the shape object for the image. Read-only.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.|
|[Line](/javascript/api/excel/excel.line)|[connectorType](/javascript/api/excel/excel.line#connectortype)|Represents the connector type for the line.|
||[id](/javascript/api/excel/excel.line#id)|Represents the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.line#shape)|Returns the shape object for the line. Read-only.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[source](/javascript/api/excel/excel.listdatavalidation#source)|Source of the list for data validation|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Deletes a page break object.|
||[getStartCell()](/javascript/api/excel/excel.pagebreak#getstartcell--)|Gets the first cell after the page break.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Represents the column index for the page break|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Represents the row index for the page break|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Adds a page break before the top-left cell of the range specified.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Gets the number of page breaks in the collection.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Gets a page break object via the index.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.pagebreakcollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
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
||[setPrintMargins(unit: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Sets the worksheet's page margins with units.|
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
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortby: "Ascending" \| "Descending", valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when|
||[sortByValues(sortby: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|True if formatting will be automatically formatted when it’s refreshed or when fields are moved|
||[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|True if the field list should be shown or hidden from the UI.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Gets the cell in the PivotTable's data body that contains the value for the intersection of the specified dataHierarchy, rowItems, and columnItems.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|
||[getPivotItems(axis: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|True if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
||[setAutosortOnCell(cell: Range \| string, sortby: "Ascending" \| "Descending")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Sets an autosort using the specified cell to automatically select all criteria and context for the sort.|
||[setAutosortOnCell(cell: Range \| string, sortby: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Sets an autosort using the specified cell to automatically select all criteria and context for the sort.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|True if the PivotTable should use custom lists when sorting.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|True if the PivotTable should use custom lists when sorting.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)||
||[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)||
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Converts the range cells with datatypes into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Converts the range cells into linked datatype in the worksheet.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copies cell data or formatting from the source range or RangeAreas to the current range.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copies cell data or formatting from the source range or RangeAreas to the current range.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Finds the given string based on the criteria specified.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Finds the given string based on the criteria specified.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Returns a 2D array, encapsulating the data for each cell's font, fill, borders, alignment, and other properties.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Returns a single-dimensional array, encapsulating the data for each column's font, fill, borders, alignment, and other properties.  For properties that are not consistent across each cell within a given column, null will be returned.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Returns a single-dimensional array, encapsulating the data for each row's font, fill, borders, alignment, and other properties.  For properties that are not consistent across each cell within a given row, null will be returned.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Comments" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Comments" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Gets the range object containing the anchor cell for a cell getting spilled into. Fails if applied to a range with more than one cell. Read only.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Gets the range object containing the spill range when called on an anchor cell. Fails if applied to a range with more than one cell. Read only.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Gets a scoped collection of tables that overlap with the range.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Represents if all cells have a spill border.|
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
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Comments" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Comments" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.|
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
||[track()](/javascript/api/excel/excel.rangeareas#track--)|Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Returns or sets a double that lightens or darkens a color for Range Borders, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Returns the number of ranges in the RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Returns the range object based on its position in the RangeCollection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.rangecollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
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
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)||
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)||
||[hasSpill](/javascript/api/excel/excel.rowproperties#hasspill)||
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)||
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Specifies whether the match needs to be complete or partial. Default is false (partial).|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Specifies whether the match is case sensitive. Default is false (insensitive).|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Specifies the search direction. Default is forward. See Excel.SearchDirection.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)||
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)||
||[style](/javascript/api/excel/excel.settablecellproperties#style)||
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)||
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[format: Excel.CellPropertiesFormat & {
            columnWidth?](/javascript/api/excel/excel.settablecolumnproperties#format)||
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel.CellPropertiesFormat & {
            rowHeight?](/javascript/api/excel/excel.settablerowproperties#format)||
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)||
|[Setting](/javascript/api/excel/excel.setting)|[](/javascript/api/excel/excel.setting#replacestringdatewithdate)||
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Returns or sets the alternative descriptive text string for a Shape object when the object is saved to a Web page.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Returns or sets the alternative title text string for a Shape object when the object is saved to a Web page.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Deletes the Shape|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Represents the geometric shape type of the specified shape. See Excel.GeometricShapeType for detail. Returns null if the shape is not geometric, for example, get GeometricShapeType of a line or a chart will return null.|
||[height](/javascript/api/excel/excel.shape#height)|Represents the height, in points, of the shape.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Moves the shape horizontally by the specified number of points.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|Changes the rotation of the shape around the z-axis by the specified number of degrees.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Moves the shape vertically by the specified number of points.|
||[left](/javascript/api/excel/excel.shape#left)|The distance, in points, from the left side of the shape to the left of the worksheet.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Represents if the aspect ratio locked, in boolean, of the shape.|
||[name](/javascript/api/excel/excel.shape#name)|Represents the name of the shape.|
||[placement](/javascript/api/excel/excel.shape#placement)|Represents the placment, value that represents the way the object is attached to the cells below it.|
||[fill](/javascript/api/excel/excel.shape#fill)|Returns the fill formatting of the shape object. Read-only.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Returns the geometric shape for the shape object. Error will be thrown, if the shape object is other shape type (Like, Image, SmartArt, etc.) rather than GeometricShape.|
||[group](/javascript/api/excel/excel.shape#group)|Returns the shape group for the shape object. Error will be thrown, if the shape object is other shape type (Like, Image, SmartArt, etc.) rather than GroupShape.|
||[id](/javascript/api/excel/excel.shape#id)|Represents the shape identifier. Read-only.|
||[image](/javascript/api/excel/excel.shape#image)|Returns the image for the shape object. Error will be thrown, if the shape object is other shape type (Like, GeometricShape, SmartArt, etc.) rather than Image.|
||[level](/javascript/api/excel/excel.shape#level)|Represents the level of the specified shape. Level 0 means the shape is not part of any group, level 1 means the shape is part of a top-level group, etc.|
||[line](/javascript/api/excel/excel.shape#line)|Returns the line object for the shape object. Error will be thrown, if the shape object is other shape type (Like, GeometricShape, SmartArt, etc.) rather than Image.|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Returns the line formatting of the shape object. Read-only.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Occurs when the shape is activated.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Occurs when the shape is activated.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Represents the parent group of the specified shape.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Returns the textFrame object of a shape. Read only.|
||[type](/javascript/api/excel/excel.shape#type)|Returns the type of the specified shape. Read-only. See Excel.ShapeType for detail.|
||[zorderPosition](/javascript/api/excel/excel.shape#zorderposition)|Returns the position of the specified shape in the z-order, the very bottom shape's z-order value is 0. Read-only.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Represents the rotation, in degrees, of the shape.|
||[saveAsPicture(format: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#saveaspicture-format-)|Saves the shape as a picture and returns the picture in the form of base64 encoded string, using the DPI sets to 96. Only support saves as to Excel.PictureFormat.BMP, Excel.PictureFormat.PNG, Excel.PictureFormat.JPEG and Excel.PictureFormat.GIF.|
||[saveAsPicture(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#saveaspicture-format-)|Saves the shape as a picture and returns the picture in the form of base64 encoded string, using the DPI sets to 96. Only support saves as to Excel.PictureFormat.BMP, Excel.PictureFormat.PNG, Excel.PictureFormat.JPEG and Excel.PictureFormat.GIF.|
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Scales the height of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Scales the height of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.|
||[setZOrder(value: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-value-)|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|
||[setZOrder(value: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-value-)|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|
||[top](/javascript/api/excel/excel.shape#top)|The distance, in points, from the top edge of the shape to the top of the worksheet.|
||[visible](/javascript/api/excel/excel.shape#visible)|Represents the visibility, in boolean, of the specified shape.|
||[width](/javascript/api/excel/excel.shape#width)|Represents the width, in points, of the shape.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Gets the id of the shape that is activated.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Gets the id of the worksheet in which the shape is activated.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus", left: number, top: number, width: number, height: number)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype--left--top--width--height-)|Adds a geometric shape to worksheet. Returns a Shape object that represents the new shape.|
||[addGeometricShape(geometricShapeType: Excel.GeometricShapeType, left: number, top: number, width: number, height: number)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype--left--top--width--height-)|Adds a geometric shape to worksheet. Returns a Shape object that represents the new shape.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Group a subset of shapes in a worksheet. Returns a Shape object that represents the new group of shapes.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Creates an image from a base64 string and adds it to worksheet. Returns the Shape object that represents the new Image.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adds a line to worksheet. Returns a Shape object that represents the new line.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adds a line to worksheet. Returns a Shape object that represents the new line.|
||[addSVG(xmlImageString: string)](/javascript/api/excel/excel.shapecollection#addsvg-xmlimagestring-)|Creates an SVG from a XML string and adds it to worksheet. Returns a Shape object that represents the new Image.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Adds a textbox to worksheet by telling it's text content. Returns a Shape object that represents the new text box.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Returns the number of shapes in the worksheet. Read-only.|
||[getItem(name: string)](/javascript/api/excel/excel.shapecollection#getitem-name-)|Gets a shape using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Gets a shape based on its position in the collection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.shapecollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Gets the loaded child items in this collection.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Gets the id of the shape that is deactivated.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Gets the id of the worksheet in which the shape is deactivated.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Clears the fill formatting of a shape object.|
||[foreColor](/javascript/api/excel/excel.shapefill#forecolor)|Represents the shape fill fore color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|
||[type](/javascript/api/excel/excel.shapefill#type)|Returns the fill type of the shape. Read-only. See Excel.ShapeFillType for detail.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Sets the fill formatting of a shape object to a uniform color, fill type changeing to Solid Fill.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). For API not supported shape types  or special fill type with inconsistent transparencies, return null. For example, gradient fill type could have inconsistent transparencies.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|
||[color](/javascript/api/excel/excel.shapefont#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red. Returns null if the TextRange includes text fragments with different colors.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Represents the italic status of font. Return null if the TextRange includes both italic and non-italic text fragments.|
||[name](/javascript/api/excel/excel.shapefont#name)|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, represents corresponding font name; otherwise represents Latin font name.|
||[size](/javascript/api/excel/excel.shapefont#size)|Represents font size in points (e.g. 11). Return null if the TextRange includes text fragments with different font sizes.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Type of underline applied to the font. Return null if the TextRange includes text fragments with different underline styles. See Excel.ShapeFontUnderlineStyle for details.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Represents the shape identifier. Read-only.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Returns the shape object for the group. Read-only.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Returns the shape collection in the group. Read-only.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Ungroups any grouped shapes in the specified shape group.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Represents the line style of the shape. Returns null when line is not visible or has mixed line dash style property (e.g. group type of shape). See Excel.ShapeLineStyle for details.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Represents the line style of the shape object. Returns null when line is not visible or has mixed line visible property (e.g. group type of shape). See Excel.ShapeLineStyle for details.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has mixed line transparency property (e.g. group type of shape).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Represents whether the line formatting of a shape element is visible. Returns null when the shape has mixed line visible property (e.g. group type of shape).|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Represents weight of the line, in points. Returns null when the line is not visible or has mixed line weight property (e.g. group type of shape).|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Represents the caption of slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Clears all the filters currently applied on the slicer.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Deletes the slicer.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Returns an array of selected items' names. Read-only.|
||[height](/javascript/api/excel/excel.slicer#height)|Represents the height, in points, of the slicer.|
||[left](/javascript/api/excel/excel.slicer#left)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicer#name)|Represents the name of slicer.|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Represents the name used in the formula.|
||[id](/javascript/api/excel/excel.slicer#id)|Represents the unique id of slicer. Read-only.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|True if all filters currently applied on the slicer is cleared.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Represents the collection of SlicerItems that are part of the slicer. Read-only.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Represents the worksheet containing the slicer. Read-only.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Select slicer items based on their names. Previous selection will be cleared.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Represents the sort order of the items in the slicer.|
||[style](/javascript/api/excel/excel.slicer#style)|Constant value that represents the Slicer style. Possible values are: SlicerStyleLight1 thru SlicerStyleLight6, TableStyleOther1 thru TableStyleOther2, SlicerStyleDark1 thru SlicerStyleDark6. A custom user-defined style present in the workbook can also be specified.|
||[top](/javascript/api/excel/excel.slicer#top)|Represents the distance, in points, from the top edge of the slicer to the right of the worksheet.|
||[width](/javascript/api/excel/excel.slicer#width)|Represents the width, in points, of the slicer.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Adds a new slicer to the workbook.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Returns the number of slicers in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Gets a slicer object using its name or id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Gets a slicer based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Gets a slicer using its name or id. If the slicer does not exist, will return a null object.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.slicercollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Gets the loaded child items in this collection.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True if the slicer item is selected. Setting this value will not clear other SlicerItems' selected state.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True if the slicer item has data.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Represents the value displayed on UI.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Returns the number of slicer items in the slicer.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Gets a slicer item object using its key or name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Gets a slicer item based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Gets a slicer item using its key or name. If the slicer item does not exist, will return a null object.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.sliceritemcollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Gets the loaded child items in this collection.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Represents the subfield that is the target property name of a rich value to sort on.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Gets the number of styles in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Gets a style based on its position in the collection.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Changes the table to use the default table style.|
||[autoFilter](/javascript/api/excel/excel.table#autofilter)|Represents the AutoFilter object of the table. Read-Only.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Occurs when filter is applied on a specific table.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Gets the id of the table that is added.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Gets the id of the worksheet in which the table is added.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Occurs when new table is added in a workbook.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Occurs when the specified table is deleted in a workbook.|
||[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Occurs when filter is applied on any table in a workbook, or a worksheet.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Specifies the source of the event. See Excel.EventSource for details.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Specifies the id of the table that is deleted.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Specifies the name of the table that is deleted.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Specifies the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Specifies the id of the worksheet in which the table is deleted.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Represents the id of the table in which the filter is applied..|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Represents the id of the worksheet which contains the table.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Gets the number of tables in the collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Gets the first table in the collection. The tables in the collection are sorted top to bottom and left to right, such that top left table is the first table in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Gets a table by Name or ID.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.tablescopedcollection#load-option-)|Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Gets the loaded child items in this collection.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSize](/javascript/api/excel/excel.textframe#autosize)|Gets or sets the auto sizing settings for the text frame. A text frame can be set to auto size the text to fit the text frame, or auto size the text frame to fit the text, or without auto sizing.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Deletes all the text in the textframe.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Represents the horizontal alignment of the text frame.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Represents the horizontal overflow type of the text frame.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Represents the left margin, in points, of the text frame.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Represents the text orientation of the text frame.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Represents the reading order of the text frame, RTL or LTR.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Specifies whether the TextFrame contains text.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)||
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Represents the right margin, in points, of the text frame.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Represents the vertical alignment of the text frame.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Represents the vertical overflow type of the text frame.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getCharacters(start: number, length?: number)](/javascript/api/excel/excel.textrange#getcharacters-start--length-)|Returns a TextRange object for characters in the given range.|
||[font](/javascript/api/excel/excel.textrange#font)|Returns a ShapeFont object that represents the font attributes for the text range. Read-only.|
||[text](/javascript/api/excel/excel.textrange#text)|Represents the plain text content of the text range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|Close current workbook.|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Close current workbook.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Gets the currently active chart in the workbook. If there is no active chart, will throw exception when invoke this statement|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Gets the currently active chart in the workbook. If there is no active chart, will return null object|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Gets the currently active slicer in the workbook. If there is no active slicer, will throw exception when invoke this statement.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Gets the currently active slicer in the workbook. If there is no active slicer, will return null object|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|True if the workbook is being edited by multiple users (co-authoring).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Gets the currently selected one or more ranges from the workbook. Unlike getSelectedRange(), this method returns a RangeAreas object that represents all the selected ranges.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|True if no changes have been made to the specified workbook since it was last saved.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|True if the workbook is in auto save mode.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Returns a number about the version of Excel Calculation Engine. Read-Only.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Represents a collection of Comments associated with the workbook. Read-only.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Occurs when AutoSave setting is changed on the workbook.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|True if the workbook has ever been saved locally or online.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Represents a collection of Slicers associated with the workbook. Read-only.|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|Save current workbook.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Save current workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True if the workbook uses the 1904 date system.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Gets or sets the enableCalculation property of the worksheet.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Gets the RangeAreas object, representing one or more blocks of rectangular ranges, specified by the address or name.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Represents the AutoFilter object of the worksheet. Read-Only.|
||[comments](/javascript/api/excel/excel.worksheet#comments)|Returns a collection of all the Comments objects on the worksheet. Read-only.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Occurs when filter is applied on a specific worksheet.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Occurs when format changed on a specific worksheet.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Gets the PageLayout object of the worksheet.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Returns the collection of all the Shape objects on the worksheet. Read-only.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Returns collection of slicers that are part of the worksheet. Read-only.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Finds and replaces the given string based on the criteria specified within the current worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Occurs when any worksheet in the workbook is changed.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Occurs when any worksheet's filter is applied in the workbook.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Occurs when any worksheet in the workbook has format changed.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Occurs when the selection changes on any worksheet.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Represents the id of the worksheet in which the filter is applied.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Gets the range that represents the changed area of a specific worksheet. It might return null object.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Specifies whether the match needs to be complete or partial. Default is false (partial).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Specifies whether the match is case sensitive. Default is false (insensitive).|

## What’s new in Excel JavaScript API 1.8

The Excel JavaScript API requirement set 1.8 features include APIs for PivotTables, data validation, charts, events for charts, performance options, and workbook creation.

### PivotTable

Wave 2 of the PivotTable APIs lets add-ins set the hierarchies of a PivotTable. You can now control the data and how it is aggregated. Our [PivotTable article](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables) has more on the new PivotTable functionality.

### Data Validation

Data validation gives you control of what a user enters in a worksheet. You can limit cells to pre-defined answer sets or give pop-up warnings about undesirable input. Learn more about [adding data validation to ranges](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation) today.

### Charts

Another round of Chart APIs brings even greater programmatic control over chart elements. You now have greater access to the legend, axes, trendline, and plot area.

### Events

More [events](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events) have been added for charts. Have your add-in react to users interacting with the chart. You can also [toggle events](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events) firing across the entire workbook.

|Object| What's new| Description|Requirement Set|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Method_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|Creates a new hidden workbook by using an optional base64 encoded .xlsx file.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Property_ > formula2|Gets or sets the Formula2, i.e. maximum value or value depending of the operator.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Relationship_ > operator|The operator to use for validating the data.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > categoryLabelLevel|Returns or sets a ChartCategoryLabelLevel enumeration constant referring to the level of where the category labels are being sourced from. Read/Write.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > plotVisibleOnly|True if only visible cells are plotted. False if both visible and hidden cells are plotted. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > seriesNameLevel|Returns or sets a ChartSeriesNameLevel enumeration constant referring to the level of where the series names are being sourced from. Read/Write.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > showDataLabelsOverMaximum|Represents whether to show the data labels when the value is greater than the maximum value on the value axis.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > style|Returns or sets the chart style for the chart. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relationship_ > displayBlanksAs|Returns or sets the way that blank cells are plotted on a chart. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relationship_ > plotArea|Represents the plotArea for the chart. Read-only.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relationship_ > plotBy|Returns or sets the way columns or rows are used as data series on the chart. ReadWrite.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Property_ > chartId|Gets the id of the chart that is activated.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Property_ > type|Gets the type of the event.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the chart is activated.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Property_ > chartId|Gets the id of the chart that is added to the worksheet.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Property_ > type|Gets the type of the event.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the chart is added.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Relationship_ > source|Gets the source of the event.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > isBetweenCategories|Represents whether value axis crosses the category axis between categories.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > multiLevel|Represents whether an axis is multilevel or not.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > numberFormat|Represents the format code for the axis tick label.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > offset|Represents the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > positionAt|Represents the specified axis position where the other axis crosses at. You should use the SetPositionAt(double) method to set this property. Read-only.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > textOrientation|Represents the text orientation of the axis tick label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > alignment|Represents the alignment for the specified axis tick label.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > position|Represents the specified axis position where the other axis crosses.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Method_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|Set the specified axis position where the other axis crosses at.|1.8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_Relationship_ > fill|Represents chart fill formatting. Read-only.|1.8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_Method_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|A string value that represents the formula of chart axis title using A1-style notation.|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relationship_ > border|Represents the border format, which includes color, linestyle, and weight. Read-only.|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relationship_ > fill|Represents chart fill formatting. Read-only.|1.8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Method_ > [clear()](/javascript/api/excel/excel.chartborder)|Clear the border format of a chart element.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > autoText|Boolean value representing if data label automatically generates appropriate text based on context.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > formula|String value that represents the formula of chart data label using A1-style notation.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > height|Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible. Read-only.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > left|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > numberFormat|String value that represents the format code for data label.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > text|String representing the text of the data label on a chart.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > textOrientation|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > top|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > width|Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible. Read-only.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relationship_ > format|Represents the format of chart data label. Read-only.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment for chart data label.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relationship_ > verticalAlignment|Represents the vertical alignment of chart data label.|1.8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_Relationship_ > border|Represents the border format, which includes color, linestyle, and weight. Read-only.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Property_ > autoText|Represents whether data labels automatically generate appropriate text based on context.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Property_ > numberFormat|Represents the format code for data labels.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Property_ > textOrientation|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 0 to 180 for vertically-oriented text.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment for chart data label.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relationship_ > verticalAlignment|Represents the vertical alignment of chart data label.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Property_ > chartId|Gets the id of the chart that is deactivated.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Property_ > type|Gets the type of the event.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the chart is deactivated.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Property_ > chartId|Gets the id of the chart that is deleted from the worksheet.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Property_ > type|Gets the type of the event.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the chart is deleted.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Relationship_ > source|Gets the source of the event.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > height|Represents the height of the legendEntry on the chart legend. Read-only.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > index|Represents the index of the legendEntry in the chart legend. Read-only.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > left|Represents the left of a chart legendEntry. Read-only.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > top|Represents the top of a chart legendEntry. Read-only.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > width|Represents the width of the legendEntry on the chart Legend. Read-only.|1.8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_Relationship_ > border|Represents the border format, which includes color, linestyle, and weight. Read-only.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > height|Represents the height value of plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > insideHeight|Represents the insideHeight value of plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > insideLeft|Represents the insideLeft value of plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > insideTop|Represents the insideTop value of plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > insideWidth|Represents the insideWidth value of plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > left|Represents the left value of plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > top|Represents the top value of plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > width|Represents the width value of plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relationship_ > format|Represents the formatting of a chart plotArea. Read-only.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relationship_ > position|Represents the position of plotArea.|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relationship_ > border|Represents the border attributes of a chart plotArea. Read-only.|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relationship_ > fill|Represents the fill format of an object, which includes background formatting information. Read-only.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > explosion|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > firstSliceAngle|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. ReadWrite|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > invertIfNegative|True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > overlap|Specifies how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2-D bar and 2-D column charts. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > secondPlotSize|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > varyByCategories|True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > axisGroup|Returns or sets the group for the specified series. ReadWrite|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > dataLabels|Represents a collection of all dataLabels in the series. Read-only.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > splitType|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. ReadWrite.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > backwardPeriod|Represents the number of periods that the trendline extends backward.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > forwardPeriod|Represents the number of periods that the trendline extends forward.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > showEquation|True if the equation for the trendline is displayed on the chart.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > showRSquared|True if the R-squared for the trendline is displayed on the chart.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relationship_ > label|Represents the label of a chart trendline. Read-only.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > autoText|Boolean value representing if trendline label automatically generates appropriate text based on context.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > formula|String value that represents the formula of chart trendline label using A1-style notation.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > height|Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible. Read-only.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > left|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > numberFormat|String value that represents the format code for trendline label.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > text|String representing the text of the trendline label on a chart.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > textOrientation|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > top|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > width|Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible. Read-only.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relationship_ > format|Represents the format of chart trendline label. Read-only.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment for chart trendline label.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relationship_ > verticalAlignment|Represents the vertical alignment of chart trendline label.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relationship_ > border|Represents the border format, which includes color, linestyle, and weight. Read-only.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relationship_ > fill|Represents the fill format of the current chart trendline label. Read-only.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relationship_ > font|Represents the font attributes (font name, font size, color, etc.) for a chart trendline label. Read-only.|1.8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_Property_ > formula| A custom data validation formula. This creates special input rules, such as preventing duplicates or limiting the total in a range of cells.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Property_ > id|Id of the DataPivotHierarchy. Read-only.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Property_ > name|Name of the DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Property_ > numberFormat|Number format of the DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Property_ > position|Position of the DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relationship_ > field|Returns the PivotFields associated with the DataPivotHierarchy. Read-only.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relationship_ > showAs|Determines whether the data should be shown as a specific summary calculation or not.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relationship_ > summarizeBy|Determines whether to show all items of the DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Method_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Reset the DataPivotHierarchy back to its default values.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Property_ > items|A collection of dataPivotHierarchy objects. Read-only.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Adds the PivotHierarchy to the current axis.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|Gets the number of pivot hierarchies in the collection.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Gets a DataPivotHierarchy by its name or id.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Gets a DataPivotHierarchy by name. If the DataPivotHierarchy does not exist, will return a null object.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Removes the PivotHierarchy from the current axis.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Property_ > ignoreBlanks|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Property_ > valid|Represents if all cell values are valid according to the data validation rules. Read-only.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relationship_ > errorAlert|Error alert when user enters invalid data.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relationship_ > prompt|Prompt when users selects a cell.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relationship_ > rule|Data validation rule that contains different types of data validation criteria.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relationship_ > type|Type of the data validation, see [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) for details. Read-only.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Method_ > [clear()](/javascript/api/excel/excel.datavalidation)|Clears the data validation from the current range.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Property_ > message|Represents error alert message.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Property_ > showAlert|Determines whether to show an error alert dialog or not when a user enters invalid data. The default is true.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Property_ > title|Represents error alert dialog title.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Relationship_ > style|Represents data validation alert type, please see [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) for details.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Property_ > message|Represents the message of the prompt.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Property_ > showPrompt|Determines whether or not to show the prompt when user selects a cell with data validation.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Property_ > title|Represents the title for the prompt.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > custom|Custom data validation criteria.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > date|Date data validation criteria.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > decimal|Decimal data validation criteria.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > list|List data validation criteria.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > textLength|TextLength data validation criteria.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > time|Time data validation criteria.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > wholeNumber|WholeNumber data validation criteria.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending on the operator.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Property_ > formula2|Gets or sets the Formula2, i.e. maximum value or value depending on the operator.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Relationship_ > operator|The operator to use for validating the data.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Property_ > enableMultipleFilterItems|Determines whether to allow multiple filter items.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Property_ > id|Id of the FilterPivotHierarchy. Read-only.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Property_ > name|Name of the FilterPivotHierarchy.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Property_ > position|Position of the FilterPivotHierarchy.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Relationship_ > fields|Returns the PivotFields associated with the FilterPivotHierarchy. Read-only.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Method_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|Reset the FilterPivotHierarchy back to its default values.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Property_ > items|A collection of filterPivotHierarchy objects. Read-only.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column, or filter axis, it will be removed from that location.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|Gets the number of pivot hierarchies in the collection.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Gets a FilterPivotHierarchy by its name or id.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Gets a FilterPivotHierarchy by name. If the FilterPivotHierarchy does not exist, will return a null object.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Removes the PivotHierarchy from the current axis.|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Property_ > inCellDropDown|Displays the list in cell drop down or not, it defaults to true.|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Property_ > source|Source of the list for data validation|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Property_ > id|Id of the PivotField. Read-only.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Property_ > name|Name of the PivotField.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Property_ > showAllItems|Determines whether to show all items of the PivotField.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relationship_ > items|Returns the PivotFields associated with the PivotField. Read-only.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relationship_ > subtotals|Subtotals of the PivotField.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Method_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Property_ > items|A collection of pivotField objects. Read-only.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Method_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|Gets the number of pivot hierarchies in the collection.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Gets a PivotHierarchy by its name or id.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Property_ > id|Id of the PivotHierarchy. Read-only.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Property_ > name|Name of the PivotHierarchy.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Relationship_ > fields|Returns the PivotFields associated with the PivotHierarchy. Read-only.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Property_ > items|A collection of pivotHierarchy objects. Read-only.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Method_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|Gets the number of pivot hierarchies in the collection.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Gets a PivotHierarchy by its name or id.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Property_ > id|Id of the PivotItem. Read-only.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Property_ > isExpanded|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Property_ > name|Name of the PivotItem.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Property_ > visible|Determines whether the PivotItem is visible or not.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Property_ > items|A collection of pivotItem objects. Read-only.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Method_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|Gets the number of pivot hierarchies in the collection.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Gets a PivotHierarchy by its name or id.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Property_ > showColumnGrandTotals|True if the PivotTable report shows grand totals for columns.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Property_ > showRowGrandTotals|True if the PivotTable report shows grand totals for rows.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Property_ > subtotalLocation|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null. Possible values are: AtTop, AtBottom.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Relationship_ > layoutType|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|Returns the range where the PivotTable's column labels reside.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|Returns the range where the PivotTable's data values reside.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|Returns the range of the PivotTable's filter area.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|Returns the range the PivotTable exists on, excluding the filter area.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|Returns the range where the PivotTable's row labels reside.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > columnHierarchies|The Column Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > dataHierarchies|The Data Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > filterHierarchies|The Filter Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > hierarchies|The Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > layout|The PivotLayout describing the layout and visual structure of the PivotTable. Read-only.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > rowHierarchies|The Row Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Method_ > [delete()](/javascript/api/excel/excel.pivottable)|Deletes the PivotTable.|1.8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|Add a Pivottable based on the specified source data and insert it at the top left cell of the destination range.|1.8|
|[range](/javascript/api/excel/excel.range)|_Relationship_ > dataValidation|Returns a data validation object. Read-only.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Property_ > id|Id of the RowColumnPivotHierarchy. Read-only.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Property_ > name|Name of the RowColumnPivotHierarchy.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Property_ > position|Position of the RowColumnPivotHierarchy.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Relationship_ > fields|Returns the PivotFields associated with the RowColumnPivotHierarchy. Read-only.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Method_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|Reset the RowColumnPivotHierarchy back to its default values.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Property_ > items|A collection of rowColumnPivotHierarchy objects. Read-only.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Gets the number of pivot hierarchies in the collection.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Gets a RowColumnPivotHierarchy by its name or id.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, will return a null object.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Removes the PivotHierarchy from the current axis.|1.8|
|[runtime](/javascript/api/excel/excel.runtime)|_Property_ > enableEvents|Toggle JavaScript events in the current taskpane or content add-in.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relationship_ > baseField|The base PivotField to base the ShowAs calculation, if applicable based on the ShowAsCalculation type, else null.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relationship_ > baseItem|The base Item to base the ShowAs calculation on, if applicable based on the ShowAsCalculation type, else null.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relationship_ > calculation|The ShowAs Calculation to use for the Data PivotField.|1.8|
|[style](/javascript/api/excel/excel.style)|_Property_ > autoIndent|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|1.8|
|[style](/javascript/api/excel/excel.style)|_Property_ > textOrientation|The text orientation for the style.|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > automatic|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > average| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > count| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > countNumbers| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > max| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > min| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > product| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > standardDeviation| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > standardDeviationP| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > sum| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > variance| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > varianceP| |1.8|
|[table](/javascript/api/excel/excel.table)|_Property_ > legacyId|Returns a numeric id. Read-only.|1.8|
|[workbook](/javascript/api/excel/excel.workbook)|_Property_ > readOnly|True if the workbook is open in Read-only mode. Read-only.|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Property_ > id|Returns a value that uniquely identifies the WorkbookCreated object. Read-only.|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Method_ > [open()](/javascript/api/excel/excel.workbookcreated)|Open the workbook.|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > showGridlines|Gets or sets the worksheet's gridlines flag.|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > showHeadings|Gets or sets the worksheet's headings flag.|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Property_ > type|Gets the type of the event.|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is calculated.|1.8|

## What's new in Excel JavaScript API 1.7

The Excel JavaScript API requirement set 1.7 features include APIs for charts, events, worksheets, ranges, document properties, named items, protection options and styles.

### Customize charts

With the new chart APIs, you can create additional chart types, add a data series to a chart, set the chart title, add an axis title, add display unit, add a trendline with moving average, change a trendline to linear, and more. The following are some examples:

* Chart axis - get, set, format and remove axis unit, label and title in a chart.
* Chart series - add, set, and delete a series in a chart.  Change series markers, plot orders and sizing.
* Chart trendlines - add, get, and format trendlines in a chart.
* Chart legend - format the legend font in a chart.
* Chart point - set chart point color.
* Chart title substring -  get and set title substring for a chart.
* Chart type - option to create more chart types.

### Events

Excel events APIs provide a variety of event handlers that allow your add-in to automatically run a designated function when a specific event occurs. You can design that function to perform whatever actions your scenario requires. For a list of events that are currently available, see [Work with Events using the Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### Customize the appearance of worksheets and ranges

Using the new APIs, you can customize the appearance of worksheets in multiple ways:

* Freeze panes to keep specific rows or columns visible when you scroll in the worksheet. For example, if the first row in your worksheet contains headers, you might freeze that row so that the column headers will remain visible as you scroll down the worksheet.
* Modify the worksheet tab color.
* Add worksheet headings.


You can customize the appearance of ranges in multiple ways:

* Set the cell style for a range to ensure sure that all cells in the range have consistent formatting. A cell style is a defined set of formatting characteristics, such as fonts and font sizes, number formats, cell borders, and cell shading. Use any of Excel's built-in cell styles or create your own custom cell style.
* Set the text orientation for a range.
* Add or modify a hyperlink on a range that links to another location in the workbook or to an external location.

### Manage document properties

Using the document properties APIs, you can access built-in document properties and also create and manage custom document properties to store state of the workbook and drive workflow and business logic.

### Copy worksheets

Using the worksheet copy APIs, you can copy the data and format from one worksheet to a new worksheet within the same workbook and reduce the amount of data transfer needed.

### Handle ranges with ease

Using the various range APIs, you can do things such as get the surrounding region, get a resized range, and more. These APIs should make tasks like range manipulation and addressing much more efficient.

In addition:

* Workbook and worksheet protection options - use these APIs to protect data in a worksheet and the workbook structure.
* Update a named item - use this API to update a named item.
* Get active cell  - use this API to get the active cell of a workbook.

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > chartType|Represents the type of the chart. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > id|The unique id of chart. Read-only.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > showAllFieldButtons|Represents whether to display all field buttons on a PivotChart.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Relationship_ > border|Represents the border format of chart area, which includes color, linestyle and weight. Read-only.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Method_ > getItem(type: string, group: string)|Returns the specific axis identified by type and group.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > axisBetweenCategories|Represents whether value axis crosses the category axis between categories.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > axisGroup|Represents the group for the specified axis. Read-only. Possible values are: Primary, Secondary.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > categoryType|Returns or sets the category axis type. Possible values are: Automatic, TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > crosses|Represents the specified axis where the other axis crosses. Possible values are: Automatic, Maximum, Minimum, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > crossesAt|Represents the specified axis where the other axis crosses at. Read Only. Set to this property should use SetCrossesAt(double) method. Read-only.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > customDisplayUnit|Represents the custom axis display unit value. Read Only. To set this property, please use the SetCustomDisplayUnit(double) method. Read-only.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > displayUnit|Represents the axis display unit. Possible values are: None, Hundreds, Thousands, TenThousands, HundredThousands, Millions, TenMillions, HundredMillions, Billions, Trillions, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > height|Represents the height, in points, of the chart axis. Null if the axis's not visible. Read-only.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > left|Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis's not visible. Read-only.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > logBase|Represents the base of the logarithm when using logarithmic scales.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > reversePlotOrder|Represents whether Microsoft Excel plots data points from last to first.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > scaleType|Represents the value axis scale type. Possible values are: Linear, Logarithmic.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > showDisplayUnitLabel|Represents whether the axis display unit label is visible.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > tickLabelSpacing|Represents the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > tickMarkSpacing|Represents the number of categories or series between tick marks.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > top|Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis's not visible. Read-only.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > type|Represents the axis type. Read-only. Possible values are: Invalid, Category, Value, Series.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > visible|A boolean value represents the visibility of the axis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > width|Represents the width, in points, of the chart axis. Null if the axis's not visible. Read-only.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > baseTimeUnit|Returns or sets the base unit for the specified category axis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > majorTickMark|Represents the type of major tick mark for the specified axis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > majorTimeUnitScale|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > minorTickMark|Represents the type of minor tick mark for the specified axis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > minorTimeUnitScale|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > tickLabelPosition|Represents the position of tick-mark labels on the specified axis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Method_ > setCategoryNames(sourceData: Range)|Sets all the category names for the specified axis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Method_ > setCrossesAt(value: double)|Set the specified axis where the other axis crosses at.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Method_ > setCustomDisplayUnit(value: double)|Sets the axis display unit to a custom value.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Property_ > color|HTML color code representing the color of borders in the chart.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Property_ > weight|Represents weight of the border, in points.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Relationship_ > lineStyle|Represents the line style of the border.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > position|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > separator|String representing the separator used for the data label on a chart.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showBubbleSize|Boolean value representing if the data label bubble size is visible or not.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showCategoryName|Boolean value representing if the data label category name is visible or not.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showLegendKey|Boolean value representing if the data label legend key is visible or not.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showPercentage|Boolean value representing if the data label percentage is visible or not.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showSeriesName|Boolean value representing if the data label series name is visible or not.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showValue|Boolean value representing if the data label value is visible or not.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > height|Represents the height of the legend on the chart.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > left|Represents the left of a chart legend.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > showShadow|Represents if the legend has shadow on the chart.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > top|Represents the top of a chart legend.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > width|Represents the width of the legend on the chart.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Relationship_ > legendEntries|Represents a collection of legendEntries in the legend. Read-only.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > visible|Represents the visible of a chart legend entry.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Property_ > items|A collection of chartLegendEntry objects. Read-only.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Method_ > getCount()|Returns the number of legendEntry in the collection.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Method_ > getItemAt(index: number)|Returns a legendEntry at the given index.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > hasDataLabel|Represents whether a data point has datalabel. Not applicable for surface charts.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > markerBackgroundColor|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > markerForegroundColor|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > markerSize|Represents marker size of data point.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > markerStyle|Represents marker style of a chart data point. Possible values are: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Relationship_ > dataLabel|Returns the data label of a chart point. Read-only.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Relationship_ > border|Represents the border format of a chart data point, which includes color, style and weight information. Read-only.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > chartType|Represents the chart type of a series. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > doughnutHoleSize|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > filtered|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > gapWidth|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > hasDataLabels|Boolean value representing if the series has data labels or not.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > markerBackgroundColor|Represents markers background color of a chart series.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > markerForegroundColor|Represents markers foreground color of a chart series.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > markerSize|Represents marker size of a chart series.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > markerStyle|Represents marker style of a chart series. Possible values are: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > plotOrder|Represents the plot order of a chart series within the chart group.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > showShadow|Boolean value representing if the series has shadow or not.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > smooth|Boolean value representing if the series is smooth or not. Only for line and scatter charts.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > dataLabels|Represents a collection of all dataLabels in the series. Read-only.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > trendlines|Represents a collection of trendlines in the series. Read-only.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Method_ > delete()|Deletes the chart series.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Method_ > setBubbleSizes(sourceData: Range)|Set bubble sizes for a chart series. Only works for bubble charts.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Method_ > setValues(sourceData: Range)|Set values for a chart series. For scatter chart, it means Y axis values.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Method_ > setXAxisValues(sourceData: Range)|Set values of X axis for a chart series. Only works for scatter charts.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Method_ > add(name: string, index: number)|Add a new series to the collection.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > height|Returns the height, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > horizontalAlignment|Represents the horizontal alignment for chart title. Possible values are: Center, Left, Justify, Distributed, Right.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > left|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title's not visible.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > position|Represents the position of chart title. Possible values are: Top, Automatic, Bottom, Right, Left.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > showShadow|Represents a boolean value that determines if the chart title has a shadow.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > textOrientation|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > top|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title's not visible.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > verticalAlignment|Represents the vertical alignment of chart title. Possible values are: Center, Bottom, Top, Justify, Distributed.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > width|Returns the width, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Method_ > setFormula(formula: string)|Sets a string value that represents the formula of chart title using A1-style notation.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Relationship_ > border|Represents the border format of chart title, which includes color, linestyle and weight. Read-only.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > backward|Represents the number of periods that the trendline extends backward.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > displayEquation|True if the equation for the trendline is displayed on the chart.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > displayRSquared|True if the R-squared for the trendline is displayed on the chart.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > forward|Represents the number of periods that the trendline extends forward.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > intercept|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > movingAveragePeriod|Represents the period of a chart trendline, only for trendline with MovingAverage type.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > name|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > polynomialOrder|Represents the order of a chart trendline, only for trendline with Polynomial type.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > type|Represents the type of a chart trendline. Possible values are: Linear, Exponential, Logarithmic, MovingAverage, Polynomial, Power.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relationship_ > format|Represents the formatting of a chart trendline. Read-only.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Method_ > delete()|Delete the trendline object.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Property_ > items|A collection of chartTrendline objects. Read-only.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Method_ > add(type: string)|Adds a new trendline to trendline collection.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Method_ > getCount()|Returns the number of trendlines in the collection.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Method_ > getItem(index: number)|Get trendline object by index, which is the insertion order in items array.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Relationship_ > line|Represents chart line formatting. Read-only.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Property_ > key|Gets the key of the custom property. Read only. Read-only.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Property_ > type|Gets the value type of the custom property. Read only. Read-only. Possible values are: Number, Boolean, Date, String, Float.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Property_ > value|Gets or sets the value of the custom property.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Method_ > delete()|Deletes the custom property.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Property_ > items|A collection of customProperty objects. Read-only.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > add(key: string, value: object)|Creates a new or sets an existing custom property.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > deleteAll()|Deletes all custom properties in this collection.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > getCount()|Gets the count of custom properties.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > getItem(key: string)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > getItemOrNullObject(key: string)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Property_ > items|A collection of dataConnection objects. Read-only.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Method_ > refreshAll()|Refreshes all the Data Connections in the collection.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > author|Gets or sets the author of the workbook.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > category|Gets or sets the category of the workbook.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > comments|Gets or sets the comments of the workbook.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > company|Gets or sets the company of the workbook.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > keywords|Gets or sets the keywords of the workbook.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > lastAuthor|Gets the last author of the workbook. Read only. Read-only.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > manager|Gets or sets the manager of the workbook.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > revisionNumber|Gets the revision number of the workbook. Read only.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > subject|Gets or sets the subject of the workbook.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > title|Gets or sets the title of the workbook.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relationship_ > creationDate|Gets the creation date of the workbook. Read only. Read-only.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relationship_ > custom|Gets the collection of custom properties of the workbook. Read only. Read-only.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Property_ > formula|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relationship_ > arrayValues|Returns an object containing values and types of the named item. Read-only.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Property_ > types|Represents the types for each item in the named item array Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Property_ > values|Represents the values of each item in the named item array. Read-only.|1.7|
|[range](/javascript/api/excel/excel.range)|_Property_ > isEntireColumn|Represents if the current range is an entire column. Read-only.|1.7|
|[range](/javascript/api/excel/excel.range)|_Property_ > isEntireRow|Represents if the current range is an entire row. Read-only.|1.7|
|[range](/javascript/api/excel/excel.range)|_Property_ > numberFormatLocal|Represents Excel's number format code for the given range as a string in the language of the user.|1.7|
|[range](/javascript/api/excel/excel.range)|_Property_ > style|Represents the style of the current range. This return either null or a string.|1.7|
|[range](/javascript/api/excel/excel.range)|_Method_ > getAbsoluteResizedRange(numRows: number, numColumns: number)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|1.7|
|[range](/javascript/api/excel/excel.range)|_Method_ > getImage()|Renders the range as a base64-encoded image.|1.7|
|[range](/javascript/api/excel/excel.range)|_Method_ > getSurroundingRegion()|Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|1.7|
|[range](/javascript/api/excel/excel.range)|_Method_ > showCard()|Displays the card for an active cell if it has rich value content.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > textOrientation|Gets or sets the text orientation of all the cells within the range.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > useStandardHeight|Determines if the row height of the Range object equals the standard height of the sheet.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > useStandardWidth|Determines if the columnwidth of the Range object equals the standard width of the sheet.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Property_ > address|Represents the url target for the hyperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Property_ > document..|Represents the document .. target for the hyperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Property_ > screenTip|Represents the string displayed when hovering over the hyperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Property_ > textToDisplay|Represents the string that is displayed in the top left most cell in the range.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > addIndent|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > autoIndent|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > builtIn|Indicates if the style is a built-in style. Read-only.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > formulaHidden|Indicates if the formula will be hidden when the worksheet is protected.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > horizontalAlignment|Represents the horizontal alignment for the style. Possible values are: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeAlignment|Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeBorder|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeFont|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeNumber|Indicates if the style includes the NumberFormat property.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includePatterns|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeProtection|Indicates if the style includes the FormulaHidden and Locked protection properties.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > indentLevel|An integer from 0 to 250 that indicates the indent level for the style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > locked|Indicates if the object is locked when the worksheet is protected.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > name|The name of the style. Read-only.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > numberFormat|The format code of the number format for the style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > numberFormatLocal|The localized format code of the number format for the style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > orientation|The text orientation for the style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > readingOrder|The reading order for the style. Possible values are: Context, LeftToRight, RightToLeft.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > shrinkToFit|Indicates if text automatically shrinks to fit in the available column width.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > textOrientation|The text orientation for the style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > verticalAlignment|Represents the vertical alignment for the style. Possible values are: Top, Center, Bottom, Justify, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > wrapText|Indicates if Microsoft Excel wraps the text in the object.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relationship_ > borders|A Border collection of four Border objects that represent the style of the four borders. Read-only.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relationship_ > fill|The Fill of the style. Read-only.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relationship_ > font|A Font object that represents the font of the style. Read-only.|1.7|
|[style](/javascript/api/excel/excel.style)|_Method_ > delete()|Deletes this style.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Property_ > items|A collection of style objects. Read-only.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Method_ > add(name: string)]|Adds a new style to the collection.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Method_ > getItem(name: string)|Gets a style by name.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > address|Gets the address that represents the changed area of a table on a specific worksheet.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > changeType|Gets the change type that represents how the Changed event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > tableId|Gets the id of the table in which the data changed.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > address|Gets the range address that represents the selected area of the table on a specific worksheet.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > isInsideTable|Indicates if the selection is inside a table, address will be useless if IsInsideTable is false.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > tableId|Gets the id of the table in which the selection changed.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Property_ > name|Gets the workbook name. Read-only.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > dataConnections|Refreshes all data connections in the workbook. Read-only.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > properties|Gets the workbook properties. Read-only.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > protection|Returns workbook protection object for a workbook. Read-only.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > styles|Represents a collection of styles associated with the workbook. Read-only.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Method_ > getActiveCell()|Gets the currently active cell from the workbook.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Property_ > protected|Indicates if the workbook is protected. Read-Only. Read-only.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Method_ > protect(password: string)|Protects a workbook. Fails if the workbook has been protected.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Method_ > unprotect(password: string)|Unprotects a workbook.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > gridlines|Gets or sets the worksheet's gridlines flag.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > headings|Gets or sets the worksheet's headings flag.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > showHeadings|Gets or sets the worksheet's headings flag.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > standardHeight|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > standardWidth|Returns or sets the standard (default) width of all the columns in the worksheet.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > tabColor|Gets or sets the worksheet tab color.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relationship_ > freezePanes|Gets an object that can be used to manipulate frozen panes on the worksheet Read-only.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|Copy a worksheet and place it at the specified position. Return the copied worksheet.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is activated.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is added to the workbook.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > address|Gets the range address that represents the changed area of a specific worksheet.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > changeType|Gets the change type that represents how the Changed event is triggered. Possible values are: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is deactivated.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Property_ > source|Gets the source of the event. Possible values are: Local, Remote.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet that is deleted from the workbook.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > freezeAt(frozenRange: Range or string)|Sets the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > freezeColumns(count: number)|Freeze the first column(s) of the worksheet in place.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > freezeRows(count: number)|Freeze the top row(s) of the worksheet in place.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > getLocation()|Gets a range that describes the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > getLocationOrNullObject()|Gets a range that describes the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > unfreeze()|Removes all frozen panes in the worksheet.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowEditObjects|Represents the worksheet protection option of allowing editing objects.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowEditScenarios|Represents the worksheet protection option of allowing editing scenarios.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Relationship_ > selectionMode|Represents the worksheet protection option of selection mode.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > address|Gets the range address that represents the selected area of a specific worksheet.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > type|Gets the type of the event. Possible values are: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|1.7|


## What's new in Excel JavaScript API 1.6 

### Conditional formatting

Introduces conditional formating of a range. Allows the following types of conditional formatting:

* Color scale
* Data bar
* Icon set
* Custom

In addition:

* Returns the range the conditional format is applied to. 
* Removal of conditional formatting. 
* Provides priority and stopifTrue capability. 
* Get collection of all conditional formatting on a given range. 
* Clears all conditional formats active on the current specified range. 

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Method_ > suspendApiCalculationUntilNextSync()|Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relationship_ > rule|Represents the Rule object on this conditional format.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Property_ > threeColorScale|If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum). Read-only.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Relationship_ > criteria|The criteria of the color scale. Midpoint is optional when using a two point color scale.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Property_ > formula1|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Property_ > formula2|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Property_ > operator|The operator of the text conditional format. Possible values are: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > maximum|The maximum point Color Scale Criterion.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > midpoint|The midpoint Color Scale Criterion if the color scale is a 3-color scale.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > minimum|The minimum point Color Scale Criterion.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Property_ > color|HTML color code representation of the color scale color. E.g. #FF0000 represents Red.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Property_ > formula|A number, a formula, or null (if Type is LowestValue).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Property_ > type|What the icon conditional formula should be based on. Possible values are: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > borderColor|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > fillColor|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > matchPositiveBorderColor|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > matchPositiveFillColor|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > borderColor|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > fillColor|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > gradientFill|Boolean representation of whether or not the DataBar has a gradient.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Property_ > formula|The formula, if required, to evaluate the databar rule on.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Property_ > type|The type of rule for the databar. Possible values are: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Property_ > id|The Priority of the Conditional Format within the current ConditionalFormatCollection. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Property_ > priority|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Property_ > stopIfTrue|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Property_ > type|A type of conditional format. Only one can be set at a time. Read-Only. Read-only. Possible values are: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > cellValue|Returns the cell value conditional format properties if the current conditional format is a CellValue type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > cellValueOrNullObject|Returns the cell value conditional format properties if the current conditional format is a CellValue type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > colorScale|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > colorScaleOrNullObject|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > custom|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > customOrNullObject|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > dataBar|Returns the data bar properties if the current conditional format is a data bar. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > dataBarOrNullObject|Returns the data bar properties if the current conditional format is a data bar. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > iconSet|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > iconSetOrNullObject|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > preset|Returns the preset criteria conditional format such as above averagebelow averageunique valuescontains blanknonblankerrornoerror properties. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > presetOrNullObject|Returns the preset criteria conditional format such as above averagebelow averageunique valuescontains blanknonblankerrornoerror properties. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > textComparison|Returns the specific text conditional format properties if the current conditional format is a text type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > textComparisonOrNullObject|Returns the specific text conditional format properties if the current conditional format is a text type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > topBottom|Returns the TopBottom conditional format properties if the current conditional format is an TopBottom type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > topBottomOrNullObject|Returns the TopBottom conditional format properties if the current conditional format is an TopBottom type. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Method_ > delete()|Deletes this conditional format.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Method_ > getRange()|Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Method_ > getRangeOrNullObject()|Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Property_ > items|A collection of conditionalFormat objects. Read-only.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > add(type: string)|Adds a new conditional format to the collection at the firsttop priority.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > clearAll()|Clears all conditional formats active on the current specified range.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > getCount()|Returns the number of conditional formats in the workbook. Read-only.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > getItem(id: string)|Returns a conditional format for the given ID.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > getItemAt(index: number)|Returns a conditional format at the given index.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Property_ > formula|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Property_ > formulaLocal|The formula, if required, to evaluate the conditional format rule on in the user's language.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Property_ > formulaR1C1|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Property_ > formula|A number or a formula depending on the type.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Property_ > operator|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format. Possible values are: Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relationship_ > customIcon|The custom icon for the current criterion if different from the default IconSet, else null will be returned.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relationship_ > type|What the icon conditional formula should be based on.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Property_ > criterion|The criterion of the conditional format. Possible values are: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Property_ > color|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Property_ > id|Represents border identifier. Read-only. Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Property_ > sideIndex|Constant value that indicates the specific side of the border. Read-only. Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Property_ > style|One of the constants of line style specifying the line style for the border. Possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Property_ > count|Number of border objects in the collection. Read-only.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Property_ > items|A collection of conditionalRangeBorder objects. Read-only.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > bottom|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > left|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > right|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > top|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Method_ > getItem(index: string)|Gets a border object using its name|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Method_ > getItemAt(index: number)|Gets a border object using its index|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Property_ > color|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Method_ > clear()|Resets the fill.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > bold|Represents the bold status of font.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > color|HTML color code representation of the text color. E.g. #FF0000 represents Red.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > italic|Represents the italic status of the font.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > strikethrough|Represents the strikethrough status of the font.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > underline|Type of underline applied to the font. Possible values are: None, Single, Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Method_ > clear()|Resets the font formats.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Property_ > numberFormat|Represents Excel's number format code for the given range. Cleared if null is passed in.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relationship_ > borders|Collection of border objects that apply to the overall conditional format range. Read-only.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relationship_ > fill|Returns the fill object defined on the overall conditional format range. Read-only.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relationship_ > font|Returns the font object defined on the overall conditional format range. Read-only.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Property_ > operator|The operator of the text conditional format. Possible values are: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Property_ > text|The Text value of conditional format.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Property_ > rank|The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Property_ > type|Format values based on the top or bottom rank. Possible values are: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relationship_ > rule|Represents the Rule object on this conditional format. Read-only.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Property_ > axisColor|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Property_ > axisFormat|Representation of how the axis is determined for an Excel data bar. Possible values are: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Property_ > barDirection|Represents the direction that the data bar graphic should be based on. Possible values are: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Property_ > showDataBarOnly|If true, hides the values from the cells where the data bar is applied.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relationship_ > lowerBoundRule|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relationship_ > negativeFormat|Representation of all values to the left of the axis in an Excel data bar. Read-only.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relationship_ > positiveFormat|Representation of all values to the right of the axis in an Excel data bar. Read-only.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relationship_ > upperBoundRule|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Property_ > reverseIconOrder|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Property_ > showIconOnly|If true, hides the values and only shows icons.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Property_ > style|If set, displays the IconSet option for the conditional format. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Relationship_ > criteria|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula and operator will be ignored when set.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relationship_ > rule|The rule of the conditional format.|1.6|
|[range](/javascript/api/excel/excel.range)|_Relationship_ > conditionalFormats|Collection of ConditionalFormats that intersect the range. Read-only.|1.6|
|[range](/javascript/api/excel/excel.range)|_Method_ > calculate()|Calculates a range of cells on a worksheet.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relationship_ > rule|The rule of the conditional format.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relationship_ > rule|The criteria of the TopBottom conditional format.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > internalTest|For internal use only. Read-only.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > calculate(markAllDirty: bool)|Calculates all cells on a worksheet.|1.6|

##  What's new in Excel JavaScript API 1.5

### Custom XML part

* Addition of custom XML parts collection to workbook object.
* Get custom XML part using ID
* Get a new scoped collection of custom XML parts whose namespaces match the given namespace.
* Get XML string associated with a part.
* Provide id and namespace of a part.
* Adds a new custom XML part to the workbook.
* Set entire XML part.
* Delete a custom XML part.
* Delete an attribute with the given name from the element identified by xpath.
* Query the XML content by xpath.
* Insert, update and delete attribute.

**Reference implementation:** Please refer [here](https://github.com/mandren/Excel-CustomXMLPart-Demo) for a reference implementation that shows how custom XML parts can be used in an add-in.

### Others
* `range.getSurroundingRegion()` Returns a Range object that represents the surrounding region for this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.
* `getNextColumn()` and `getPreviousColumn()`, `getLast() on table column.
* `getActiveWorksheet()` on the workbook.
* `getRange(address: string)` off of workbook.
* `getBoundingRange(ranges: )` Gets the smallest range object that encompasses the provided ranges. For example, the bounding range between "B2:C5" and "D10:E15" is "B2:E15".
* `getCount()` on various collections such as named item, worksheet, table, etc. to get number of items in a collection. `workbook.worksheets.getCount()`
* `getFirst()` and `getLast()` and get last on various collection such as tworksheet, able column, chart points, range view collection.
* `getNext()` and `getPrevious()` on worksheet, table column collection.
* `getRangeR1C1()` Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Property_ > id|The custom XML part's ID. Read-only.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Property_ > namespaceUri|The custom XML part's namespace URI. Read-only.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Method_ > delete()|Deletes the custom XML part.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Method_ > getXml()|Gets the custom XML part's full XML content.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Method_ > setXml(xml: string)|Sets the custom XML part's full XML content.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Property_ > items|A collection of customXmlPart objects. Read-only.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > add(xml: string)|Adds a new custom XML part to the workbook.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > getByNamespace(namespaceUri: string)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > getCount()|Gets the number of CustomXml parts in the collection.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > getItem(id: string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > getItemOrNullObject(id: string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Property_ > items|A collection of customXmlPartScoped objects. Read-only.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getCount()|Gets the number of CustomXML parts in this collection.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getItem(id: string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getItemOrNullObject(id: string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getOnlyItem()|If the collection contains exactly one item, this method returns it.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getOnlyItemOrNullObject()|If the collection contains exactly one item, this method returns it.|1.5|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > customXmlParts|Represents the collection of custom XML parts contained by this workbook. Read-only.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getNext(visibleOnly: bool)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getNextOrNullObject(visibleOnly: bool)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getPrevious(visibleOnly: bool)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getPreviousOrNullObject(visibleOnly: bool)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Method_ > getFirst(visibleOnly: bool)|Gets the first worksheet in the collection.|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Method_ > getLast(visibleOnly: bool)|Gets the last worksheet in the collection.|1.5|

## What's new in Excel JavaScript API 1.4
The following are the new additions to the Excel JavaScript APIs in requirement set 1.4.

### Named item add and new properties

New properties:

* `comment`
* `scope` worksheet or workbook scoped items
* `worksheet` returns the worksheet on which the named item is scoped to.

New methods:

* `add(name: string, reference: Range or string, comment: string)`Adds a new name to the collection of the given scope.
* `addFormulaLocal(name: string, formula: string, comment: string)` Adds a new name to the collection of the given scope using the user's locale for the formula.

### Settings API in the Excel namespace

The [Setting](/javascript/api/excel/excel.setting) object represents a key:value pair for a setting persisted to the document. The functionality of `Excel.Setting` is equivalent to `Office.Settings`, but uses the batched API syntax, rather than the Common API's callback model.

APIs include `getItem()` to get setting entry via the key, `add()` to add the specified key:value setting pair to the workbook.

### Others

* Set table column name (prior version only allows reading).
* Add table column to the end of the table (prior version only allows anywhere but last).
* Add multiple rows to a table at a time (prior version only allows 1 row at a time).
* `range.getColumnsAfter(count: number)` and `range.getColumnsBefore(count: number)` to get a certain number of columns to the right/left of the current Range object.
* Get item or null object function: This functionality allows getting object using a key. If the object does not exist, the returned object's isNullObject property will be true. This alows developers to check if an object exists or not without having to handle it thorugh exception handling. Available on worksheet, named-item, binding, chart series, etc.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > getCount()|Gets the number of bindings in the collection.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > getItemOrNullObject(id: string)|Gets a binding object by ID. If the binding object does not exist, will return a null object.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Method_ > getCount()|Returns the number of charts in the worksheet.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Method_ > getItemOrNullObject(name: string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Method_ > getCount()|Returns the number of chart points in the series.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Method_ > getCount()|Returns the number of series in the collection.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Property_ > comment|Represents the comment associated with this name.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Property_ > scope|Indicates whether the name is scoped to the workbook or to a specific worksheet. Read-only. Possible values are: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relationship_ > worksheet|Returns the worksheet on which the named item is scoped to. Throws an error if the items is scoped to the workbook instead. Read-only.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relationship_ > worksheetOrNullObject|Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead. Read-only.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Method_ > delete()|Deletes the given name.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Method_ > getRangeOrNullObject()|Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > add(name: string, reference: Range or string, comment: string)|Adds a new name to the collection of the given scope.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > addFormulaLocal(name: string, formula: string, comment: string)|Adds a new name to the collection of the given scope using the user's locale for the formula.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > getCount()|Gets the number of named items in the collection.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > getItemOrNullObject(name: string)|Gets a nameditem object using its name. If the nameditem object does not exist, will return a null object.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > getCount()|Gets the number of pivot tables in the collection.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > getItemOrNullObject(name: string)|Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.|1.4|
|[range](/javascript/api/excel/excel.range)|_Method_ > getIntersectionOrNullObject(anotherRange: Range or string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.4|
|[range](/javascript/api/excel/excel.range)|_Method_ > getUsedRangeOrNullObject(valuesOnly: bool)|Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Method_ > getCount()|Gets the number of RangeView objects in the collection.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Property_ > value|Represents the value stored for this setting.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Method_ > delete()|Deletes the setting.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Property_ > items|A collection of setting objects. Read-only.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > add(key: string, value: (any))|Sets or adds the specified setting to the workbook.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getCount()|Gets the number of Settings in the collection.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getItem(key: string)|Gets a Setting entry via the key.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getItemOrNullObject(key: string)|Gets a Setting entry via the key. If the Setting does not exist, will return a null object.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relationship_ > settings|Gets the Setting object that represents the binding that raised the SettingsChanged event|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Method_ > getCount()]|Gets the number of tables in the collection.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Method_ > getItemOrNullObject(key: number or string)|Gets a table by Name or ID. If the table does not exist, will return a null object.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Method_ > getCount()|Gets the number of columns in the table.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Method_ > getItemOrNullObject(key: number or string)|Gets a column object by Name or ID. If the column does not exist, will return a null object.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Method_ > getCount()|Gets the number of rows in the table.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relationship_ > names|Collection of names scoped to the current worksheet. Read-only.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getUsedRangeOrNullObject(valuesOnly: bool)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Method_ > getCount(visibleOnly: bool)|Gets the number of worksheets in the collection.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Method_ > getItemOrNullObject(key: string)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.|1.4|

## What's new in Excel JavaScript API 1.3

The following are the new additions to the Excel JavaScript APIs in requirement set 1.3.

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Method_ > delete()|Deletes the binding.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > add(range: Range or string, bindingType: string, id: string)|Add a new binding to a particular Range.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > addFromNamedItem(name: string, bindingType: string, id: string)|Add a new binding based on a named item in the workbook.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > addFromSelection(bindingType: string, id: string)|Add a new binding based on the current selection.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > getItemOrNull(id: string)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Method_ > getItemOrNull(name: string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > getItemOrNull(name: string)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Property_ > name|Name of the PivotTable.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > worksheet|The worksheet containing the current PivotTable. Read-only.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Method_ > refresh()|Refreshes the PivotTable.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Property_ > items|A collection of pivotTable objects. Read-only.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > getItem(name: string)|Gets a PivotTable by name.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > getItemOrNull(name: string)|Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.|1.3|
|[range](/javascript/api/excel/excel.range)|_Method_ > getIntersectionOrNull(anotherRange: Range or string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.3|
|[range](/javascript/api/excel/excel.range)|_Method_ > getVisibleView()|Represents the visible rows of the current range.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > cellAddresses|Represents the cell addresses of the RangeView. Read-only.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > columnCount|Returns the number of visible columns. Read-only.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > formulas|Represents the formula in A1-style notation.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > formulasLocal|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, introduced in 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > index|Returns a value that represents the index of the RangeView. Read-only.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > numberFormat|Represents Excel's number format code for the given cell.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > rowCount|Returns the number of visible rows. Read-only.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > text|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > valueTypes|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > values|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Relationship_ > rows|Represents a collection of range views associated with the range. Read-only.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Method_ > getRange()|Gets the parent range associated with the current RangeView.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Property_ > items|A collection of rangeView objects. Read-only.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Method_ > getItemAt(index: number)|Gets a RangeView Row via it's index. Zero-Indexed.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Method_ > delete()|Deletes the setting.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Property_ > items|A collection of setting objects. Read-only.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getItem(key: string)|Gets a Setting entry via the key.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getItemOrNull(key: string)|Gets a Setting entry via the key. If the Setting does not exist, the returned object's isNull property will be true.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > set(key: string, value: string)|Sets or adds the specified setting to the workbook.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relationship_ > settingCollection|Gets the Setting object that represents the binding that raised the SettingsChanged event|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > highlightFirstColumn|Indicates whether the first column contains special formatting.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > highlightLastColumn|Indicates whether the last column contains special formatting.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > showBandedColumns|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > showBandedRows|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > showFilterButton|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Method_ > getItemOrNull(key: number or string)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Method_ > getItemOrNull(key: number or string)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > pivotTables|Represents a collection of PivotTables associated with the workbook. Read-only.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relationship_ > pivotTables|Collection of PivotTables that are part of the worksheet. Read-only.|1.3|

## What's new in Excel JavaScript API 1.2

The following are the new additions to the Excel JavaScript APIs in requirement set 1.2.

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > id|Gets a chart based on its position in the collection. Read-only.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Relationship_ > worksheet|The worksheet containing the current chart. Read-only.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Method_ > getImage(height: number, width: number, fittingMode: string)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Relationship_ > criteria|The currently applied filter on the given column. Read-only.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > apply(criteria: FilterCriteria)|Apply the given filter criteria on the given column.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyBottomItemsFilter(count: number)|Apply a "Bottom Item" filter to the column for the given number of elements.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyBottomPercentFilter(percent: number)]|Apply a "Bottom Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyCellColorFilter(color: string)|Apply a "Cell Color" filter to the column for the given color.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyCustomFilter(criteria1: string, criteria2: string, oper: string)|Apply a "Icon" filter to the column for the given criteria strings.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyDynamicFilter(criteria: string)|Apply a "Dynamic" filter to the column.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyFontColorFilter(color: string)|Apply a "Font Color" filter to the column for the given color.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyIconFilter(icon: Icon)|Apply a "Icon" filter to the column for the given icon.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyTopItemsFilter(count: number)|Apply a "Top Item" filter to the column for the given number of elements.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyTopPercentFilter(percent: number)|Apply a "Top Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyValuesFilter(values: ())|Apply a "Values" filter to the column for the given values.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > clear()|Clear the filter on the given column.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > color|The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > criterion1|The first criterion used to filter data. Used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > criterion2|The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > dynamicCriteria|The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering. Possible values are: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > filterOn|The property used by the filter to determine whether the values should stay visible. Possible values are: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > operator|The operator used to combine criterion 1 and 2 when using "custom" filtering. Possible values are: And, Or.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > values|The set of values to be used as part of "values" filtering.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Relationship_ > icon|The icon used to filter cells. Used with "icon" filtering.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Property_ > date|The date in ISO8601 format used to filter data.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Property_ > specificity|How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009. Possible values are: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Property_ > formulaHidden|Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Property_ > locked|Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Property_ > index|Represents the index of the icon in the given set.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Property_ > set|Represents the set that the icon is part of. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Property_ > columnHidden|Represents if all columns of the current range are hidden.|1.2|
|[range](/javascript/api/excel/excel.range)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.2|
|[range](/javascript/api/excel/excel.range)|_Property_ > hidden|Represents if all cells of the current range are hidden. Read-only.|1.2|
|[range](/javascript/api/excel/excel.range)|_Property_ > rowHidden|Represents if all rows of the current range are hidden.|1.2|
|[range](/javascript/api/excel/excel.range)|_Relationship_ > sort|Represents the range sort of the current range. Read-only.|1.2|
|[range](/javascript/api/excel/excel.range)|_Method_ > merge(across: bool)|Merge the range cells into one region in the worksheet.|1.2|
|[range](/javascript/api/excel/excel.range)|_Method_ > unmerge()|Unmerge the range cells into separate cells.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > columnWidth|Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > rowHeight|Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Relationship_ > protection|Returns the format protection object for a range. Read-only.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Method_ > autofitColumns()|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Method_ > autofitRows()|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Property_ > address|Represents the visible rows of the current range.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Method_ > apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|Perform a sort operation.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > ascending|Represents whether the sorting is done in an ascending fashion.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > color|Represents the color that is the target of the condition if the sorting is on font or cell color.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > dataOption|Represents additional sorting options for this field. Possible values are: Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > key|Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > sortOn|Represents the type of sorting of this condition. Possible values are: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Relationship_ > icon|Represents the icon that is the target of the condition if the sorting is on the cell's icon.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relationship_ > sort|Represents the sorting for the table. Read-only.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relationship_ > worksheet|The worksheet containing the current table. Read-only.|1.2|
|[table](/javascript/api/excel/excel.table)|_Method_ > clearFilters()|Clears all the filters currently applied on the table.|1.2|
|[table](/javascript/api/excel/excel.table)|_Method_ > convertToRange()|Converts the table into a normal range of cells. All data is preserved.|1.2|
|[table](/javascript/api/excel/excel.table)|_Method_ > reapplyFilters()|Reapplies all the filters currently on the table.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Relationship_ > filter|Retrieve the filter applied to the column. Read-only.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Property_ > matchCase|Represents whether the casing impacted the last sort of the table. Read-only.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Property_ > method|Represents Chinese character ordering method last used to sort the table. Read-only. Possible values are: PinYin, StrokeCount.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Relationship_ > fields|Represents the current conditions used to last sort the table. Read-only.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Method_ > apply(fields: SortField, matchCase: bool, method: string)|Perform a sort operation.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Method_ > clear()|Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Method_ > reapply()|Reapplies the current sorting parameters to the table.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > functions|Represents Excel application instance that contains this workbook. Read-only.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relationship_ > protection|Returns sheet protection object for a worksheet. Read-only.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Property_ > protected|Indicates if the worksheet is protected. Read-Only. Read-only.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Relationship_ > options|Sheet protection options. Read-only.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Method_ > protect(options: WorksheetProtectionOptions)|Protects a worksheet. Fails if the worksheet has been protected.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Method_ > unprotect()|Unprotects a worksheet.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowAutoFilter|Represents the worksheet protection option of allowing using auto filter feature.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowDeleteColumns|Represents the worksheet protection option of allowing deleting columns.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowDeleteRows|Represents the worksheet protection option of allowing deleting rows.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatCells|Represents the worksheet protection option of allowing formatting cells.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatColumns|Represents the worksheet protection option of allowing formatting columns.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatRows|Represents the worksheet protection option of allowing formatting rows.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertColumns|Represents the worksheet protection option of allowing inserting columns.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertHyperlinks|Represents the worksheet protection option of allowing inserting hyperlinks.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertRows|Represents the worksheet protection option of allowing inserting rows.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowPivotTables|Represents the worksheet protection option of allowing using PivotTable feature.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowSort|Represents the worksheet protection option of allowing using sort feature.|1.2|

## Excel JavaScript API 1.1

Excel JavaScript API 1.1 is the first version of the API. For details about the API,  see the [Excel JavaScript API](/javascript/api/excel) reference topics.

## See also

- [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
