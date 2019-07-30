---
title: Excel JavaScript API requirement set 1.6
description: 'Details about the ExcelApi 1.6 requirement set'
ms.date: 07/25/2019
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
* Provides priority and stopifTrue capability.
* Get collection of all conditional formatting on a given range.
* Clears all conditional formats active on the current specified range.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.6. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.6 or earlier, see [Excel APIs in requirement set 1.6 or earlier](/javascript/api/excel?view=excel-js-1.6).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Represents the Rule object on this conditional format.|
||[set(properties: Excel.CellValueConditionalFormat)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.CellValueConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[CellValueConditionalFormatData](/javascript/api/excel/excel.cellvalueconditionalformatdata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatdata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformatdata#rule)|Represents the Rule object on this conditional format.|
|[CellValueConditionalFormatLoadOptions](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#rule)|Represents the Rule object on this conditional format.|
|[CellValueConditionalFormatUpdateData](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#rule)|Represents the Rule object on this conditional format.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|The criteria of the color scale. Midpoint is optional when using a two point color scale.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).|
||[set(properties: Excel.ColorScaleConditionalFormat)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ColorScaleConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ColorScaleConditionalFormatData](/javascript/api/excel/excel.colorscaleconditionalformatdata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatdata#criteria)|The criteria of the color scale. Midpoint is optional when using a two point color scale.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformatdata#threecolorscale)|If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).|
|[ColorScaleConditionalFormatLoadOptions](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#criteria)|The criteria of the color scale. Midpoint is optional when using a two point color scale.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#threecolorscale)|If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).|
|[ColorScaleConditionalFormatUpdateData](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata#criteria)|The criteria of the color scale. Midpoint is optional when using a two point color scale.|
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
||[set(properties: Excel.ConditionalDataBarNegativeFormat)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ConditionalDataBarNegativeFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ConditionalDataBarNegativeFormatData](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivebordercolor)|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivefillcolor)|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|
|[ConditionalDataBarNegativeFormatLoadOptions](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivebordercolor)|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivefillcolor)|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|
|[ConditionalDataBarNegativeFormatUpdateData](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivebordercolor)|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivefillcolor)|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Boolean representation of whether or not the DataBar has a gradient.|
||[set(properties: Excel.ConditionalDataBarPositiveFormat)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ConditionalDataBarPositiveFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ConditionalDataBarPositiveFormatData](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#gradientfill)|Boolean representation of whether or not the DataBar has a gradient.|
|[ConditionalDataBarPositiveFormatLoadOptions](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#gradientfill)|Boolean representation of whether or not the DataBar has a gradient.|
|[ConditionalDataBarPositiveFormatUpdateData](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#bordercolor)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#fillcolor)|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#gradientfill)|Boolean representation of whether or not the DataBar has a gradient.|
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
||[set(properties: Excel.ConditionalFormat)](/javascript/api/excel/excel.conditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.conditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: "Custom" \| "DataBar" \| "ColorScale" \| "IconSet" \| "TopBottom" \| "PresetCriteria" \| "ContainsText" \| "CellValue")](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Adds a new conditional format to the collection at the first/top priority.|
||[add(type: Excel.ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Adds a new conditional format to the collection at the first/top priority.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Clears all conditional formats active on the current specified range.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Returns the number of conditional formats in the workbook. Read-only.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Returns a conditional format for the given ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Returns a conditional format at the given index.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Gets the loaded child items in this collection.|
|[ConditionalFormatCollectionLoadOptions](/javascript/api/excel/excel.conditionalformatcollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalue)|For EACH ITEM in the collection: Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalueornullobject)|For EACH ITEM in the collection: Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[colorScale](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscale)|For EACH ITEM in the collection: Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscaleornullobject)|For EACH ITEM in the collection: Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.|
||[custom](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#custom)|For EACH ITEM in the collection: Returns the custom conditional format properties if the current conditional format is a custom type.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#customornullobject)|For EACH ITEM in the collection: Returns the custom conditional format properties if the current conditional format is a custom type.|
||[dataBar](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databar)|For EACH ITEM in the collection: Returns the data bar properties if the current conditional format is a data bar.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databarornullobject)|For EACH ITEM in the collection: Returns the data bar properties if the current conditional format is a data bar.|
||[iconSet](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconset)|For EACH ITEM in the collection: Returns the IconSet conditional format properties if the current conditional format is an IconSet type.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconsetornullobject)|For EACH ITEM in the collection: Returns the IconSet conditional format properties if the current conditional format is an IconSet type.|
||[id](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#id)|For EACH ITEM in the collection: The Priority of the Conditional Format within the current ConditionalFormatCollection. Read-only.|
||[preset](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#preset)|For EACH ITEM in the collection: Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#presetornullobject)|For EACH ITEM in the collection: Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[priority](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#priority)|For EACH ITEM in the collection: The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#stopiftrue)|For EACH ITEM in the collection: If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
||[textComparison](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparison)|For EACH ITEM in the collection: Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparisonornullobject)|For EACH ITEM in the collection: Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottom)|For EACH ITEM in the collection: Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottomornullobject)|For EACH ITEM in the collection: Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[type](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#type)|For EACH ITEM in the collection: A type of conditional format. Only one can be set at a time. Read-only.|
|[ConditionalFormatData](/javascript/api/excel/excel.conditionalformatdata)|[cellValue](/javascript/api/excel/excel.conditionalformatdata#cellvalue)|Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatdata#cellvalueornullobject)|Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[colorScale](/javascript/api/excel/excel.conditionalformatdata#colorscale)|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatdata#colorscaleornullobject)|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|
||[custom](/javascript/api/excel/excel.conditionalformatdata#custom)|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatdata#customornullobject)|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|
||[dataBar](/javascript/api/excel/excel.conditionalformatdata#databar)|Returns the data bar properties if the current conditional format is a data bar. Read-only.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatdata#databarornullobject)|Returns the data bar properties if the current conditional format is a data bar. Read-only.|
||[iconSet](/javascript/api/excel/excel.conditionalformatdata#iconset)|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatdata#iconsetornullobject)|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|
||[id](/javascript/api/excel/excel.conditionalformatdata#id)|The Priority of the Conditional Format within the current ConditionalFormatCollection. Read-only.|
||[preset](/javascript/api/excel/excel.conditionalformatdata#preset)|Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatdata#presetornullobject)|Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[priority](/javascript/api/excel/excel.conditionalformatdata#priority)|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatdata#stopiftrue)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
||[textComparison](/javascript/api/excel/excel.conditionalformatdata#textcomparison)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatdata#textcomparisonornullobject)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/javascript/api/excel/excel.conditionalformatdata#topbottom)|Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatdata#topbottomornullobject)|Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[type](/javascript/api/excel/excel.conditionalformatdata#type)|A type of conditional format. Only one can be set at a time. Read-only.|
|[ConditionalFormatLoadOptions](/javascript/api/excel/excel.conditionalformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalue)|Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalueornullobject)|Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[colorScale](/javascript/api/excel/excel.conditionalformatloadoptions#colorscale)|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#colorscaleornullobject)|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.|
||[custom](/javascript/api/excel/excel.conditionalformatloadoptions#custom)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#customornullobject)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[dataBar](/javascript/api/excel/excel.conditionalformatloadoptions#databar)|Returns the data bar properties if the current conditional format is a data bar.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#databarornullobject)|Returns the data bar properties if the current conditional format is a data bar.|
||[iconSet](/javascript/api/excel/excel.conditionalformatloadoptions#iconset)|Returns the IconSet conditional format properties if the current conditional format is an IconSet type.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#iconsetornullobject)|Returns the IconSet conditional format properties if the current conditional format is an IconSet type.|
||[id](/javascript/api/excel/excel.conditionalformatloadoptions#id)|The Priority of the Conditional Format within the current ConditionalFormatCollection. Read-only.|
||[preset](/javascript/api/excel/excel.conditionalformatloadoptions#preset)|Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#presetornullobject)|Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[priority](/javascript/api/excel/excel.conditionalformatloadoptions#priority)|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatloadoptions#stopiftrue)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
||[textComparison](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparison)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparisonornullobject)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/javascript/api/excel/excel.conditionalformatloadoptions#topbottom)|Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#topbottomornullobject)|Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[type](/javascript/api/excel/excel.conditionalformatloadoptions#type)|A type of conditional format. Only one can be set at a time. Read-only.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|The formula, if required, to evaluate the conditional format rule on.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|The formula, if required, to evaluate the conditional format rule on in the user's language.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|
||[set(properties: Excel.ConditionalFormatRule)](/javascript/api/excel/excel.conditionalformatrule#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ConditionalFormatRuleUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.conditionalformatrule#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ConditionalFormatRuleData](/javascript/api/excel/excel.conditionalformatruledata)|[formula](/javascript/api/excel/excel.conditionalformatruledata#formula)|The formula, if required, to evaluate the conditional format rule on.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruledata#formulalocal)|The formula, if required, to evaluate the conditional format rule on in the user's language.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruledata#formular1c1)|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|
|[ConditionalFormatRuleLoadOptions](/javascript/api/excel/excel.conditionalformatruleloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatruleloadoptions#$all)||
||[formula](/javascript/api/excel/excel.conditionalformatruleloadoptions#formula)|The formula, if required, to evaluate the conditional format rule on.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleloadoptions#formulalocal)|The formula, if required, to evaluate the conditional format rule on in the user's language.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleloadoptions#formular1c1)|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|
|[ConditionalFormatRuleUpdateData](/javascript/api/excel/excel.conditionalformatruleupdatedata)|[formula](/javascript/api/excel/excel.conditionalformatruleupdatedata#formula)|The formula, if required, to evaluate the conditional format rule on.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleupdatedata#formulalocal)|The formula, if required, to evaluate the conditional format rule on in the user's language.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleupdatedata#formular1c1)|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|
|[ConditionalFormatUpdateData](/javascript/api/excel/excel.conditionalformatupdatedata)|[cellValue](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalue)|Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalueornullobject)|Returns the cell value conditional format properties if the current conditional format is a CellValue type.|
||[colorScale](/javascript/api/excel/excel.conditionalformatupdatedata#colorscale)|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#colorscaleornullobject)|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.|
||[custom](/javascript/api/excel/excel.conditionalformatupdatedata#custom)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#customornullobject)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[dataBar](/javascript/api/excel/excel.conditionalformatupdatedata#databar)|Returns the data bar properties if the current conditional format is a data bar.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#databarornullobject)|Returns the data bar properties if the current conditional format is a data bar.|
||[iconSet](/javascript/api/excel/excel.conditionalformatupdatedata#iconset)|Returns the IconSet conditional format properties if the current conditional format is an IconSet type.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#iconsetornullobject)|Returns the IconSet conditional format properties if the current conditional format is an IconSet type.|
||[preset](/javascript/api/excel/excel.conditionalformatupdatedata#preset)|Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#presetornullobject)|Returns the preset criteria conditional format. See Excel.PresetCriteriaConditionalFormat for more details.|
||[priority](/javascript/api/excel/excel.conditionalformatupdatedata#priority)|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatupdatedata#stopiftrue)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
||[textComparison](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparison)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparisonornullobject)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/javascript/api/excel/excel.conditionalformatupdatedata#topbottom)|Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#topbottomornullobject)|Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|The custom icon for the current criterion if different from the default IconSet, else null will be returned.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|A number or a formula depending on the type.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|What the icon conditional formula should be based on.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[criterion](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|The criterion of the conditional format.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Constant value that indicates the specific side of the border. See Excel.ConditionalRangeBorderIndex for details. Read-only.|
||[set(properties: Excel.ConditionalRangeBorder)](/javascript/api/excel/excel.conditionalrangeborder#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ConditionalRangeBorderUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.conditionalrangeborder#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: "EdgeTop" \| "EdgeBottom" \| "EdgeLeft" \| "EdgeRight")](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Gets a border object using its name.|
||[getItem(index: Excel.ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Gets a border object using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Gets a border object using its index.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Gets the bottom border. Read-only.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Number of border objects in the collection. Read-only.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Gets the loaded child items in this collection.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Gets the left border. Read-only.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Gets the right border. Read-only.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Gets the top border. Read-only.|
|[ConditionalRangeBorderCollectionLoadOptions](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#color)|For EACH ITEM in the collection: HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#sideindex)|For EACH ITEM in the collection: Constant value that indicates the specific side of the border. See Excel.ConditionalRangeBorderIndex for details. Read-only.|
||[style](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#style)|For EACH ITEM in the collection: One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
|[ConditionalRangeBorderCollectionUpdateData](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#bottom)|Gets the bottom border.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#left)|Gets the left border.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#right)|Gets the right border.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#top)|Gets the top border.|
|[ConditionalRangeBorderData](/javascript/api/excel/excel.conditionalrangeborderdata)|[color](/javascript/api/excel/excel.conditionalrangeborderdata#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborderdata#sideindex)|Constant value that indicates the specific side of the border. See Excel.ConditionalRangeBorderIndex for details. Read-only.|
||[style](/javascript/api/excel/excel.conditionalrangeborderdata#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
|[ConditionalRangeBorderLoadOptions](/javascript/api/excel/excel.conditionalrangeborderloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangeborderloadoptions#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborderloadoptions#sideindex)|Constant value that indicates the specific side of the border. See Excel.ConditionalRangeBorderIndex for details. Read-only.|
||[style](/javascript/api/excel/excel.conditionalrangeborderloadoptions#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
|[ConditionalRangeBorderUpdateData](/javascript/api/excel/excel.conditionalrangeborderupdatedata)|[color](/javascript/api/excel/excel.conditionalrangeborderupdatedata#color)|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[style](/javascript/api/excel/excel.conditionalrangeborderupdatedata#style)|One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Resets the fill.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[set(properties: Excel.ConditionalRangeFill)](/javascript/api/excel/excel.conditionalrangefill#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ConditionalRangeFillUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.conditionalrangefill#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ConditionalRangeFillData](/javascript/api/excel/excel.conditionalrangefilldata)|[color](/javascript/api/excel/excel.conditionalrangefilldata#color)|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
|[ConditionalRangeFillLoadOptions](/javascript/api/excel/excel.conditionalrangefillloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangefillloadoptions#color)|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
|[ConditionalRangeFillUpdateData](/javascript/api/excel/excel.conditionalrangefillupdatedata)|[color](/javascript/api/excel/excel.conditionalrangefillupdatedata#color)|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Represents the bold status of font.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Resets the font formats.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Represents the italic status of the font.|
||[set(properties: Excel.ConditionalRangeFont)](/javascript/api/excel/excel.conditionalrangefont#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ConditionalRangeFontUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.conditionalrangefont#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Represents the strikethrough status of the font.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Type of underline applied to the font. See Excel.ConditionalRangeFontUnderlineStyle for details.|
|[ConditionalRangeFontData](/javascript/api/excel/excel.conditionalrangefontdata)|[bold](/javascript/api/excel/excel.conditionalrangefontdata#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.conditionalrangefontdata#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.conditionalrangefontdata#italic)|Represents the italic status of the font.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontdata#strikethrough)|Represents the strikethrough status of the font.|
||[underline](/javascript/api/excel/excel.conditionalrangefontdata#underline)|Type of underline applied to the font. See Excel.ConditionalRangeFontUnderlineStyle for details.|
|[ConditionalRangeFontLoadOptions](/javascript/api/excel/excel.conditionalrangefontloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.conditionalrangefontloadoptions#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.conditionalrangefontloadoptions#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.conditionalrangefontloadoptions#italic)|Represents the italic status of the font.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontloadoptions#strikethrough)|Represents the strikethrough status of the font.|
||[underline](/javascript/api/excel/excel.conditionalrangefontloadoptions#underline)|Type of underline applied to the font. See Excel.ConditionalRangeFontUnderlineStyle for details.|
|[ConditionalRangeFontUpdateData](/javascript/api/excel/excel.conditionalrangefontupdatedata)|[bold](/javascript/api/excel/excel.conditionalrangefontupdatedata#bold)|Represents the bold status of font.|
||[color](/javascript/api/excel/excel.conditionalrangefontupdatedata#color)|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
||[italic](/javascript/api/excel/excel.conditionalrangefontupdatedata#italic)|Represents the italic status of the font.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontupdatedata#strikethrough)|Represents the strikethrough status of the font.|
||[underline](/javascript/api/excel/excel.conditionalrangefontupdatedata#underline)|Type of underline applied to the font. See Excel.ConditionalRangeFontUnderlineStyle for details.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Represents Excel's number format code for the given range. Cleared if null is passed in.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Collection of border objects that apply to the overall conditional format range. Read-only.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Returns the fill object defined on the overall conditional format range. Read-only.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Returns the font object defined on the overall conditional format range. Read-only.|
||[set(properties: Excel.ConditionalRangeFormat)](/javascript/api/excel/excel.conditionalrangeformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.ConditionalRangeFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.conditionalrangeformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[ConditionalRangeFormatData](/javascript/api/excel/excel.conditionalrangeformatdata)|[borders](/javascript/api/excel/excel.conditionalrangeformatdata#borders)|Collection of border objects that apply to the overall conditional format range. Read-only.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatdata#fill)|Returns the fill object defined on the overall conditional format range. Read-only.|
||[font](/javascript/api/excel/excel.conditionalrangeformatdata#font)|Returns the font object defined on the overall conditional format range. Read-only.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatdata#numberformat)|Represents Excel's number format code for the given range. Cleared if null is passed in.|
|[ConditionalRangeFormatLoadOptions](/javascript/api/excel/excel.conditionalrangeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeformatloadoptions#$all)||
||[borders](/javascript/api/excel/excel.conditionalrangeformatloadoptions#borders)|Collection of border objects that apply to the overall conditional format range.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatloadoptions#fill)|Returns the fill object defined on the overall conditional format range.|
||[font](/javascript/api/excel/excel.conditionalrangeformatloadoptions#font)|Returns the font object defined on the overall conditional format range.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatloadoptions#numberformat)|Represents Excel's number format code for the given range. Cleared if null is passed in.|
|[ConditionalRangeFormatUpdateData](/javascript/api/excel/excel.conditionalrangeformatupdatedata)|[borders](/javascript/api/excel/excel.conditionalrangeformatupdatedata#borders)|Collection of border objects that apply to the overall conditional format range.|
||[fill](/javascript/api/excel/excel.conditionalrangeformatupdatedata#fill)|Returns the fill object defined on the overall conditional format range.|
||[font](/javascript/api/excel/excel.conditionalrangeformatupdatedata#font)|Returns the font object defined on the overall conditional format range.|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatupdatedata#numberformat)|Represents Excel's number format code for the given range. Cleared if null is passed in.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|The operator of the text conditional format.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|The Text value of conditional format.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Format values based on the top or bottom rank.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.customconditionalformat#rule)|Represents the Rule object on this conditional format. Read-only.|
||[set(properties: Excel.CustomConditionalFormat)](/javascript/api/excel/excel.customconditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.CustomConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.customconditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[CustomConditionalFormatData](/javascript/api/excel/excel.customconditionalformatdata)|[format](/javascript/api/excel/excel.customconditionalformatdata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.customconditionalformatdata#rule)|Represents the Rule object on this conditional format. Read-only.|
|[CustomConditionalFormatLoadOptions](/javascript/api/excel/excel.customconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.customconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.customconditionalformatloadoptions#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.customconditionalformatloadoptions#rule)|Represents the Rule object on this conditional format.|
|[CustomConditionalFormatUpdateData](/javascript/api/excel/excel.customconditionalformatupdatedata)|[format](/javascript/api/excel/excel.customconditionalformatupdatedata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.customconditionalformatupdatedata#rule)|Represents the Rule object on this conditional format.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Representation of how the axis is determined for an Excel data bar.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Represents the direction that the data bar graphic should be based on.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Representation of all values to the left of the axis in an Excel data bar. Read-only.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Representation of all values to the right of the axis in an Excel data bar. Read-only.|
||[set(properties: Excel.DataBarConditionalFormat)](/javascript/api/excel/excel.databarconditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.DataBarConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.databarconditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|If true, hides the values from the cells where the data bar is applied.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|
|[DataBarConditionalFormatData](/javascript/api/excel/excel.databarconditionalformatdata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatdata#axiscolor)|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatdata#axisformat)|Representation of how the axis is determined for an Excel data bar.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatdata#bardirection)|Represents the direction that the data bar graphic should be based on.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatdata#lowerboundrule)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatdata#negativeformat)|Representation of all values to the left of the axis in an Excel data bar. Read-only.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatdata#positiveformat)|Representation of all values to the right of the axis in an Excel data bar. Read-only.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatdata#showdatabaronly)|If true, hides the values from the cells where the data bar is applied.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatdata#upperboundrule)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|
|[DataBarConditionalFormatLoadOptions](/javascript/api/excel/excel.databarconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.databarconditionalformatloadoptions#$all)||
||[axisColor](/javascript/api/excel/excel.databarconditionalformatloadoptions#axiscolor)|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#axisformat)|Representation of how the axis is determined for an Excel data bar.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatloadoptions#bardirection)|Represents the direction that the data bar graphic should be based on.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatloadoptions#lowerboundrule)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#negativeformat)|Representation of all values to the left of the axis in an Excel data bar.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#positiveformat)|Representation of all values to the right of the axis in an Excel data bar.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatloadoptions#showdatabaronly)|If true, hides the values from the cells where the data bar is applied.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatloadoptions#upperboundrule)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|
|[DataBarConditionalFormatUpdateData](/javascript/api/excel/excel.databarconditionalformatupdatedata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatupdatedata#axiscolor)|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#axisformat)|Representation of how the axis is determined for an Excel data bar.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatupdatedata#bardirection)|Represents the direction that the data bar graphic should be based on.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatupdatedata#lowerboundrule)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#negativeformat)|Representation of all values to the left of the axis in an Excel data bar.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#positiveformat)|Representation of all values to the right of the axis in an Excel data bar.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatupdatedata#showdatabaronly)|If true, hides the values from the cells where the data bar is applied.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatupdatedata#upperboundrule)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|
||[set(properties: Excel.IconSetConditionalFormat)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.IconSetConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|If true, hides the values and only shows icons.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|If set, displays the IconSet option for the conditional format.|
|[IconSetConditionalFormatData](/javascript/api/excel/excel.iconsetconditionalformatdata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatdata#criteria)|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatdata#reverseiconorder)|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatdata#showicononly)|If true, hides the values and only shows icons.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatdata#style)|If set, displays the IconSet option for the conditional format.|
|[IconSetConditionalFormatLoadOptions](/javascript/api/excel/excel.iconsetconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#criteria)|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#reverseiconorder)|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#showicononly)|If true, hides the values and only shows icons.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#style)|If set, displays the IconSet option for the conditional format.|
|[IconSetConditionalFormatUpdateData](/javascript/api/excel/excel.iconsetconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#criteria)|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula, and operator will be ignored when set.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#reverseiconorder)|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#showicononly)|If true, hides the values and only shows icons.|
||[style](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#style)|If set, displays the IconSet option for the conditional format.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|The rule of the conditional format.|
||[set(properties: Excel.PresetCriteriaConditionalFormat)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.PresetCriteriaConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[PresetCriteriaConditionalFormatData](/javascript/api/excel/excel.presetcriteriaconditionalformatdata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#rule)|The rule of the conditional format.|
|[PresetCriteriaConditionalFormatLoadOptions](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#rule)|The rule of the conditional format.|
|[PresetCriteriaConditionalFormatUpdateData](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#rule)|The rule of the conditional format.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Calculates a range of cells on a worksheet.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Collection of ConditionalFormats that intersect the range. Read-only.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[conditionalFormats](/javascript/api/excel/excel.rangedata#conditionalformats)|Collection of ConditionalFormats that intersect the range. Read-only.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.textconditionalformat#rule)|The rule of the conditional format.|
||[set(properties: Excel.TextConditionalFormat)](/javascript/api/excel/excel.textconditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TextConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.textconditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[TextConditionalFormatData](/javascript/api/excel/excel.textconditionalformatdata)|[format](/javascript/api/excel/excel.textconditionalformatdata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.textconditionalformatdata#rule)|The rule of the conditional format.|
|[TextConditionalFormatLoadOptions](/javascript/api/excel/excel.textconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.textconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.textconditionalformatloadoptions#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.textconditionalformatloadoptions#rule)|The rule of the conditional format.|
|[TextConditionalFormatUpdateData](/javascript/api/excel/excel.textconditionalformatupdatedata)|[format](/javascript/api/excel/excel.textconditionalformatupdatedata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.textconditionalformatupdatedata#rule)|The rule of the conditional format.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|The criteria of the Top/Bottom conditional format.|
||[set(properties: Excel.TopBottomConditionalFormat)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TopBottomConditionalFormatUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[TopBottomConditionalFormatData](/javascript/api/excel/excel.topbottomconditionalformatdata)|[format](/javascript/api/excel/excel.topbottomconditionalformatdata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformatdata#rule)|The criteria of the Top/Bottom conditional format.|
|[TopBottomConditionalFormatLoadOptions](/javascript/api/excel/excel.topbottomconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#rule)|The criteria of the Top/Bottom conditional format.|
|[TopBottomConditionalFormatUpdateData](/javascript/api/excel/excel.topbottomconditionalformatupdatedata)|[format](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#format)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#rule)|The criteria of the Top/Bottom conditional format.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Calculates all cells on a worksheet.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel&view=excel-js-1.6)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
