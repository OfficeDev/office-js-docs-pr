---
title: Excel JavaScript API requirement set 1.6
description: 'Details about the ExcelApi 1.6 requirement set'
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.6

## Conditional formatting

Introduces conditional formating of a range. Allows the following types of conditional formatting:

* Color scale
* Data bar
* Icon set
* Custom

In addition:

* Returns the range the conditional format is applied to.
* Removal of conditional formatting.
* Provides priority and `stopifTrue` capability.
* Get collection of all conditional formatting on a given range.
* Clears all conditional formats active on the current specified range.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.6. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.6 or earlier, see [Excel APIs in requirement set 1.6 or earlier](/javascript/api/excel?view=excel-js-1.6).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Represents the Rule object on this conditional format.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|The criteria of the color scale. Midpoint is optional when using a two point color scale.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|The formula, if required, to evaluate the conditional format rule on.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|The formula, if required, to evaluate the conditional format rule on.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|The operator of the text conditional format.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|The maximum point Color Scale Criterion.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|The midpoint Color Scale Criterion if the color scale is a 3-color scale.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|The minimum point Color Scale Criterion.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|HTML color code representation of the color scale color. E.g. #FF0000 represents Red.|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|A number, a formula, or null (if Type is LowestValue).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|What the criterion conditional formula should be based on.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Boolean representation of whether or not the DataBar has a gradient.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|The formula, if required, to evaluate the databar rule on.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|The type of rule for the databar.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Deletes this conditional format.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Returns the range the conditonal format is applied to. Throws an error if the conditional format is applied to multiple ranges. Read-only.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Returns the range the conditonal format is applied to, or a null object if the conditional format is applied to multiple ranges. Read-only.|
||[priority](/javascript/api/excel/excel.conditionalformat#priority)|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|
||[custom](/javascript/api/excel/excel.conditionalformat#custom)|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Returns the data bar properties if the current conditional format is a data bar. Read-only.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Returns the data bar properties if the current conditional format is a data bar. Read-only.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|The Priority of the Conditional Format within the current ConditionalFormatCollection. Read-only.|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|A type of conditional format. Only one can be set at a time. Read-only.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel.ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Adds a new conditional format to the collection at the first/top priority.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Clears all conditional formats active on the current specified range.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Returns the number of conditional formats in the workbook. Read-only.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Returns a conditional format for the given ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Returns a conditional format at the given index.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Gets the loaded child items in this collection.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|The formula, if required, to evaluate the conditional format rule on.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|The formula, if required, to evaluate the conditional format rule on in the user's language.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|The custom icon for the current criterion if different from the default IconSet, else null will be returned.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|A number or a formula depending on the type.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|What the icon conditional formula should be based on.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[criterion](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|The criterion of the conditional format.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Constant value that indicates the specific side of the border. See Excel.ConditionalRangeBorderIndex for details. Read-only.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel.ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Gets a border object using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Gets a border object using its index.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Gets the bottom border. Read-only.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Number of border objects in the collection. Read-only.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Gets the loaded child items in this collection.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Gets the left border. Read-only.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Gets the right border. Read-only.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Gets the top border. Read-only.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Resets the fill.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Represents the bold status of font.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Resets the font formats.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Represents the italic status of the font.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Represents the strikethrough status of the font.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Type of underline applied to the font. See Excel.ConditionalRangeFontUnderlineStyle for details.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Represents Excel's number format code for the given range. Cleared if null is passed in.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Collection of border objects that apply to the overall conditional format range. Read-only.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Returns the fill object defined on the overall conditional format range. Read-only.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Returns the font object defined on the overall conditional format range. Read-only.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|The operator of the text conditional format.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|The Text value of conditional format.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Format values based on the top or bottom rank.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.customconditionalformat#rule)|Represents the Rule object on this conditional format. Read-only.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Representation of how the axis is determined for an Excel data bar.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Represents the direction that the data bar graphic should be based on.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Representation of all values to the left of the axis in an Excel data bar. Read-only.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Representation of all values to the right of the axis in an Excel data bar. Read-only.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|If true, hides the values from the cells where the data bar is applied.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|If true, hides the values and only shows icons.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|If set, displays the IconSet option for the conditional format.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|The rule of the conditional format.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calculates a range of cells on a worksheet.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Collection of ConditionalFormats that intersect the range. Read-only.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.textconditionalformat#rule)|The rule of the conditional format.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|The criteria of the Top/Bottom conditional format.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calculates all cells on a worksheet.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.6)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
