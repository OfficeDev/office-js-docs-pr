---
title: Excel JavaScript API requirement set 1.6
description: 'Details about the ExcelApi 1.6 requirement set.'
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.6

## Conditional formatting

Introduces conditional formatting of a range. Allows the following types of conditional formatting.

- Color scale
- Data bar
- Icon set
- Custom

In addition:

- Returns the range the conditional format is applied to.
- Removal of conditional formatting.
- Provides priority and `stopifTrue` capability.
- Get collection of all conditional formatting on a given range.
- Clears all conditional formats active on the current specified range.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.6. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.6 or earlier, see [Excel APIs in requirement set 1.6 or earlier](/javascript/api/excel?view=excel-js-1.6&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendApiCalculationUntilNextSync__)|Suspends calculation until the next `context.sync()` is called.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Specifies the rule object on this conditional format.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|The criteria of the color scale.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threeColorScale)|If `true`, the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|The formula, if required, on which to evaluate the conditional format rule.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|The formula, if required, on which to evaluate the conditional format rule.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|The operator of the cell value conditional format.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|The maximum point of the color scale criterion.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|The midpoint of the color scale criterion, if the color scale is a 3-color scale.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|The minimum point of the color scale criterion.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|HTML color code representation of the color scale color (e.g., #FF0000 represents Red).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|A number, a formula, or `null` (if `type` is `lowestValue`).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|What the criterion conditional formula should be based on.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#borderColor)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillColor)|HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchPositiveBorderColor)|Specifies if the negative data bar has the same border color as the positive data bar.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchPositiveFillColor)|Specifies if the negative data bar has the same fill color as the positive data bar.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#borderColor)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillColor)|HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientFill)|Specifies if the data bar has a gradient.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|The formula, if required, on which to evaluate the data bar rule.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|The type of rule for the data bar.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete__)|Deletes this conditional format.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getRange__)|Returns the range the conditonal format is applied to.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getRangeOrNullObject__)|Returns the range to which the conditonal format is applied.|
||[priority](/javascript/api/excel/excel.conditionalformat#priority)|The priority (or index) within the conditional format collection that this conditional format currently exists in.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellValue)|Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellValueOrNullObject)|Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorScale)|Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorScaleOrNullObject)|Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.|
||[custom](/javascript/api/excel/excel.conditionalformat#custom)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customOrNullObject)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#dataBar)|Returns the data bar properties if the current conditional format is a data bar.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#dataBarOrNullObject)|Returns the data bar properties if the current conditional format is a data bar.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconSet)|Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconSetOrNullObject)|Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|The priority of the conditional format in the current `ConditionalFormatCollection`.|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|Returns the preset criteria conditional format.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetOrNullObject)|Returns the preset criteria conditional format.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textComparison)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textComparisonOrNullObject)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topBottom)|Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topBottomOrNullObject)|Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|A type of conditional format.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopIfTrue)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel.ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add_type_)|Adds a new conditional format to the collection at the first/top priority.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearAll__)|Clears all conditional formats active on the current specified range.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getCount__)|Returns the number of conditional formats in the workbook.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItem_id_)|Returns a conditional format for the given ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getItemAt_index_)|Returns a conditional format at the given index.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Gets the loaded child items in this collection.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|The formula, if required, on which to evaluate the conditional format rule.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulaLocal)|The formula, if required, on which to evaluate the conditional format rule in the user's language.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formulaR1C1)|The formula, if required, on which to evaluate the conditional format rule in R1C1-style notation.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customIcon)|The custom icon for the current criterion, if different from the default icon set, else `null` will be returned.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|A number or a formula depending on the type.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|`greaterThan` or `greaterThanOrEqual` for each of the rule types for the icon conditional format.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|What the icon conditional formula should be based on.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[criterion](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|The criterion of the conditional format.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideIndex)|Constant value that indicates the specific side of the border.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|One of the constants of line style specifying the line style for the border.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel.ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getItem_index_)|Gets a border object using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getItemAt_index_)|Gets a border object using its index.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Gets the bottom border.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Number of border objects in the collection.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Gets the loaded child items in this collection.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Gets the left border.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Gets the right border.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Gets the top border.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear__)|Resets the fill.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML color code representing the color of the fill, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Specifies if the font is bold.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear__)|Resets the font formats.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Specifies if the font is italic.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Specifies the strikethrough status of the font.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|The type of underline applied to the font.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberFormat)|Represents Excel's number format code for the given range.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Collection of border objects that apply to the overall conditional format range.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Returns the fill object defined on the overall conditional format range.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Returns the font object defined on the overall conditional format range.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|The operator of the text conditional format.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|The text value of the conditional format.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Format values based on the top or bottom rank.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.customconditionalformat#rule)|Specifies the `Rule` object on this conditional format.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axisColor)|HTML color code representing the color of the Axis line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisFormat)|Representation of how the axis is determined for an Excel data bar.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#barDirection)|Specifies the direction that the data bar graphic should be based on.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerBoundRule)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeFormat)|Representation of all values to the left of the axis in an Excel data bar.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveFormat)|Representation of all values to the right of the axis in an Excel data bar.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showDataBarOnly)|If `true`, hides the values from the cells where the data bar is applied.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperBoundRule)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|An array of criteria and icon sets for the rules and potential custom icons for conditional icons.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseIconOrder)|If `true`, reverses the icon orders for the icon set.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showIconOnly)|If `true`, hides the values and only shows icons.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|If set, displays the icon set option for the conditional format.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|The rule of the conditional format.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate__)|Calculates a range of cells on a worksheet.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalFormats)|The collection of `ConditionalFormats` that intersect the range.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.textconditionalformat#rule)|The rule of the conditional format.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|The criteria of the top/bottom conditional format.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate_markAllDirty_)|Calculates all cells on a worksheet.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
