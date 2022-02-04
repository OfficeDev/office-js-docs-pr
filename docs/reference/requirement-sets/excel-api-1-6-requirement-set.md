---
title: Excel JavaScript API requirement set 1.6
description: 'Details about the ExcelApi 1.6 requirement set.'
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
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
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendapicalculationuntilnextsync-member(1))|Suspends calculation until the next `context.sync()` is called.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#excel-excel-cellvalueconditionalformat-format-member)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#excel-excel-cellvalueconditionalformat-rule-member)|Specifies the rule object on this conditional format.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#excel-excel-colorscaleconditionalformat-criteria-member)|The criteria of the color scale.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#excel-excel-colorscaleconditionalformat-threecolorscale-member)|If `true`, the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-formula1-member)|The formula, if required, on which to evaluate the conditional format rule.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-formula2-member)|The formula, if required, on which to evaluate the conditional format rule.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#excel-excel-conditionalcellvaluerule-operator-member)|The operator of the cell value conditional format.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-maximum-member)|The maximum point of the color scale criterion.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-midpoint-member)|The midpoint of the color scale criterion, if the color scale is a 3-color scale.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#excel-excel-conditionalcolorscalecriteria-minimum-member)|The minimum point of the color scale criterion.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-color-member)|HTML color code representation of the color scale color (e.g., #FF0000 represents Red).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-formula-member)|A number, a formula, or `null` (if `type` is `lowestValue`).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#excel-excel-conditionalcolorscalecriterion-type-member)|What the criterion conditional formula should be based on.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-bordercolor-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-fillcolor-member)|HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-matchpositivebordercolor-member)|Specifies if the negative data bar has the same border color as the positive data bar.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#excel-excel-conditionaldatabarnegativeformat-matchpositivefillcolor-member)|Specifies if the negative data bar has the same fill color as the positive data bar.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-bordercolor-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-fillcolor-member)|HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#excel-excel-conditionaldatabarpositiveformat-gradientfill-member)|Specifies if the data bar has a gradient.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#excel-excel-conditionaldatabarrule-formula-member)|The formula, if required, on which to evaluate the data bar rule.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#excel-excel-conditionaldatabarrule-type-member)|The type of rule for the data bar.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[cellValue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalue-member)|Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalueornullobject-member)|Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscale-member)|Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscaleornullobject-member)|Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.|
||[custom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-custom-member)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-customornullobject-member)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databar-member)|Returns the data bar properties if the current conditional format is a data bar.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databarornullobject-member)|Returns the data bar properties if the current conditional format is a data bar.|
||[delete()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-delete-member(1))|Deletes this conditional format.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrange-member(1))|Returns the range the conditonal format is applied to.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrangeornullobject-member(1))|Returns the range to which the conditonal format is applied.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconset-member)|Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconsetornullobject-member)|Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.|
||[id](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-id-member)|The priority of the conditional format in the current `ConditionalFormatCollection`.|
||[preset](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-preset-member)|Returns the preset criteria conditional format.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-presetornullobject-member)|Returns the preset criteria conditional format.|
||[priority](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-priority-member)|The priority (or index) within the conditional format collection that this conditional format currently exists in.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-stopiftrue-member)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparison-member)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparisonornullobject-member)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottom-member)|Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottomornullobject-member)|Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.|
||[type](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-type-member)|A type of conditional format.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel.ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-add-member(1))|Adds a new conditional format to the collection at the first/top priority.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-clearall-member(1))|Clears all conditional formats active on the current specified range.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getcount-member(1))|Returns the number of conditional formats in the workbook.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitem-member(1))|Returns a conditional format for the given ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemat-member(1))|Returns a conditional format at the given index.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-items-member)|Gets the loaded child items in this collection.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formula-member)|The formula, if required, on which to evaluate the conditional format rule.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formulalocal-member)|The formula, if required, on which to evaluate the conditional format rule in the user's language.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#excel-excel-conditionalformatrule-formular1c1-member)|The formula, if required, on which to evaluate the conditional format rule in R1C1-style notation.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-customicon-member)|The custom icon for the current criterion, if different from the default icon set, else `null` will be returned.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-formula-member)|A number or a formula depending on the type.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-operator-member)|`greaterThan` or `greaterThanOrEqual` for each of the rule types for the icon conditional format.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#excel-excel-conditionaliconcriterion-type-member)|What the icon conditional formula should be based on.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[criterion](/javascript/api/excel/excel.conditionalpresetcriteriarule#excel-excel-conditionalpresetcriteriarule-criterion-member)|The criterion of the conditional format.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-color-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-sideindex-member)|Constant value that indicates the specific side of the border.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#excel-excel-conditionalrangeborder-style-member)|One of the constants of line style specifying the line style for the border.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-bottom-member)|Gets the bottom border.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-count-member)|Number of border objects in the collection.|
||[getItem(index: Excel.ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-getitem-member(1))|Gets a border object using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-getitemat-member(1))|Gets a border object using its index.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-items-member)|Gets the loaded child items in this collection.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-left-member)|Gets the left border.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-right-member)|Gets the right border.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#excel-excel-conditionalrangebordercollection-top-member)|Gets the top border.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#excel-excel-conditionalrangefill-clear-member(1))|Resets the fill.|
||[color](/javascript/api/excel/excel.conditionalrangefill#excel-excel-conditionalrangefill-color-member)|HTML color code representing the color of the fill, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-bold-member)|Specifies if the font is bold.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-clear-member(1))|Resets the font formats.|
||[color](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-color-member)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-italic-member)|Specifies if the font is italic.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-strikethrough-member)|Specifies the strikethrough status of the font.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#excel-excel-conditionalrangefont-underline-member)|The type of underline applied to the font.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[borders](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-borders-member)|Collection of border objects that apply to the overall conditional format range.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-fill-member)|Returns the fill object defined on the overall conditional format range.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-font-member)|Returns the font object defined on the overall conditional format range.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-numberformat-member)|Represents Excel's number format code for the given range.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#excel-excel-conditionaltextcomparisonrule-operator-member)|The operator of the text conditional format.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#excel-excel-conditionaltextcomparisonrule-text-member)|The text value of the conditional format.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#excel-excel-conditionaltopbottomrule-rank-member)|The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#excel-excel-conditionaltopbottomrule-type-member)|Format values based on the top or bottom rank.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#excel-excel-customconditionalformat-format-member)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.customconditionalformat#excel-excel-customconditionalformat-rule-member)|Specifies the `Rule` object on this conditional format.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-axiscolor-member)|HTML color code representing the color of the Axis line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-axisformat-member)|Representation of how the axis is determined for an Excel data bar.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-bardirection-member)|Specifies the direction that the data bar graphic should be based on.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-lowerboundrule-member)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-negativeformat-member)|Representation of all values to the left of the axis in an Excel data bar.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-positiveformat-member)|Representation of all values to the right of the axis in an Excel data bar.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-showdatabaronly-member)|If `true`, hides the values from the cells where the data bar is applied.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#excel-excel-databarconditionalformat-upperboundrule-member)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-criteria-member)|An array of criteria and icon sets for the rules and potential custom icons for conditional icons.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-reverseiconorder-member)|If `true`, reverses the icon orders for the icon set.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-showicononly-member)|If `true`, hides the values and only shows icons.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#excel-excel-iconsetconditionalformat-style-member)|If set, displays the icon set option for the conditional format.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#excel-excel-presetcriteriaconditionalformat-format-member)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#excel-excel-presetcriteriaconditionalformat-rule-member)|The rule of the conditional format.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#excel-excel-range-calculate-member(1))|Calculates a range of cells on a worksheet.|
||[conditionalFormats](/javascript/api/excel/excel.range#excel-excel-range-conditionalformats-member)|The collection of `ConditionalFormats` that intersect the range.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#excel-excel-textconditionalformat-format-member)|Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.textconditionalformat#excel-excel-textconditionalformat-rule-member)|The rule of the conditional format.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#excel-excel-topbottomconditionalformat-format-member)|Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#excel-excel-topbottomconditionalformat-rule-member)|The criteria of the top/bottom conditional format.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-calculate-member(1))|Calculates all cells on a worksheet.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
