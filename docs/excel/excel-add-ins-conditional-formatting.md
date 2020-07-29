---
title: Apply conditional formatting to ranges with the Excel JavaScript API
description: 'This article covers conditional formatting in the context of Excel JavaScript add-ins.'
ms.date: 07/28/2020
localization_priority: Normal
---

# Apply conditional formatting to Excel ranges

The Excel JavaScript Library provides APIs to apply conditional formatting to data ranges in your worksheets. This functionality makes large sets of data easy to visually parse. The formatting also dynamically updates based on changes within the range.

> [!NOTE]
> This article covers conditional formatting in the context of Excel JavaScript add-ins. The following articles provide detailed information about the full conditional formatting capabilities within Excel.
> -  [Add, change, or clear conditional formats](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
> -  [Use formulas with conditional formatting](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)

## Programmatic control of conditional formatting

The `Range.conditionalFormats` property is a collection of [ConditionalFormat](/javascript/api/excel/excel.conditionalformat) objects that apply to the range.  The `ConditionalFormat` object contains several properties that define the format to be applied based on the [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).

-    `cellValue`
-    `colorScale`
-    `custom`
-    `dataBar`
-    `iconSet`
-    `preset`
-    `textComparison`
-    `topBottom`

> [!NOTE]
> Each of these formatting properties has a corresponding `*OrNullObject` variant. Learn more about that pattern in the [\*OrNullObject methods](../develop/host-specific-api-model.md#ornullobject-methods-and-properties) section.

Only one format type can be set for the ConditionalFormat object. This is determined by the `type` property, which is a [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype) enum value. `type` is set when adding a conditional format to a range.

## Creating conditional formatting rules

Conditional formats are added to a range by using `conditionalFormats.add`. Once added, the properties specific to the conditional format can be set. The following examples show the creation of different formatting types.

### [Cell value](/javascript/api/excel/excel.cellvalueconditionalformat)

Cell value conditional formatting applies a user-defined format based on the results of one or two formulas in the [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule). The `operator` property is a [ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator) defining how the resulting expressions relate to the formatting.

The following example shows red font coloring applied to any value in the range less than zero.

![A range with negative numbers in red.](../images/excel-conditional-format-cell-value.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B21:E23");
const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.cellValue
);

// set the font of negative numbers to red
conditionalFormat.cellValue.format.font.color = "red";
conditionalFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };

await context.sync();
```

### [Color scale](/javascript/api/excel/excel.colorscaleconditionalformat)

Color scale conditional formatting applies a color gradient across the data range. The `criteria` property on the `ColorScaleConditionalFormat` defines three [ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`, and, optionally, `midpoint`. Each of the criterion scale points have three properties:

-    `color` - The HTML color code for the endpoint.
-    `formula` - A number or formula representing the endpoint. This will be `null` if `type` is `lowestValue` or `highestValue`.
-    `type` - How the formula should be evaluated. `highestValue` and `lowestValue` refer to values in the range being formatted.

The following example shows a range being colored blue to yellow to red. Note that `minimum` and `maximum` are the lowest and highest values respectively and use `null` formulas. `midpoint` is using the `percentage` type with a formula of `"=50"` so the yellowest cell is the mean value.

![A range with the low number in blue, average number in yellow, and high number is red, with gradients for between values.](../images/excel-conditional-format-color-scale.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
      Excel.ConditionalFormatType.colorScale
);

// color the backgrounds of the cells from blue to yellow to red based on value
const criteria = {
      minimum: {
           formula: null,
           type: Excel.ConditionalFormatColorCriterionType.lowestValue,
           color: "blue"
      },
      midpoint: {
           formula: "50",
           type: Excel.ConditionalFormatColorCriterionType.percent,
           color: "yellow"
      },
      maximum: {
           formula: null,
           type: Excel.ConditionalFormatColorCriterionType.highestValue,
           color: "red"
      }
};
conditionalFormat.colorScale.criteria = criteria;

await context.sync();
```

### [Custom](/javascript/api/excel/excel.customconditionalformat)

Custom conditional formatting applies a user-defined format to the cells based on a formula of arbitrary complexity. The [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) object lets you define the formula in different notations:

-    `formula` - Standard notation.
-    `formulaLocal` - Localized based on the user's language.
-    `formulaR1C1` - R1C1-style notation.

The following example colors the fonts green of cells with higher values than the cell to their left.

![A range with green numbers for places the preceding column's value in that row is lower.](../images/excel-conditional-format-custom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.custom
);

// if a cell has a higher value than the one to its left, set that cell's font to green
conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT("RC[-1]",0),TRUE)';
conditionalFormat.custom.format.font.color = "green";

await context.sync();

```
### [Data bar](/javascript/api/excel/excel.databarconditionalformat)

Data bar conditional formatting adds data bars to the cells. By default, the minimum and maximum values in the Range form the bounds and proportional sizes of the data bars. The `DataBarConditionalFormat` object has several properties to control the bar's appearance. 

The following example formats the range with data bars filling left-to-right.

![A range with databars behind the values in cells.](../images/excel-conditional-format-databar.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.dataBar
);

// give left-to-right, default-appearance data bars to all the cells
conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
await context.sync();
```

### [Icon set](/javascript/api/excel/excel.iconsetconditionalformat)

Icon set conditional formatting uses Excel [Icons](/javascript/api/excel/excel.icon) to highlight cells. The `criteria` property is an array of [ConditionalIconCriterion](/javascript/api/excel/excel.ConditionalIconCriterion), which define the symbol to be inserted and the condition under which it is inserted. This array is automatically prepopulated with criterion elements with default properties. Individual properties cannot be overwritten. Instead, the whole criteria object must be replaced. 

The following example shows a three-triangle icon set applied across the range.

![A range with green upward triangles for values above 1000, yellow lines for values between 700 and 1000, and red downward triangles for lower values.](../images/excel-conditional-format-iconset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B8:E13");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.iconSet
);

const iconSetCF = conditionalFormat.iconSet;
iconSetCF.style = Excel.IconSet.threeTriangles;

/*
   With a "three*" icon set style, such as "threeTriangles", the third
    element in the criteria array (criteria[2]) defines the "top" icon;
    e.g., a green triangle. The second (criteria[1]) defines the "middle"
    icon, The first (criteria[0]) defines the "low" icon, but it can often 
    be left empty as this method does below, because every cell that
   does not match the other two criteria always gets the low icon.
*/
iconSetCF.criteria = [
    {} as any,
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=700"
      },
      {
        type: Excel.ConditionalFormatIconRuleType.number,
        operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
        formula: "=1000"
      }
];

await context.sync();
```

### [Preset criteria](/javascript/api/excel/excel.presetcriteriaconditionalformat)

Preset conditional formatting applies a user-defined format to the range based on a selected standard rule. These rules are defined by the [ConditionalFormatPresetCriterion](/javascript/api/excel/excel.ConditionalFormatPresetCriterion) in the [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule). 

The following example colors the font white wherever a cell's value is at least one standard deviation above the range's average.

![A range with white font cells where the values are at least one standard deviation above average.](../images/excel-conditional-format-preset.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B2:M5");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.presetCriteria
);

// color every cell's font white that is one standard deviation above average relative to the range
conditionalFormat.preset.format.font.color = "white";
conditionalFormat.preset.rule = {
     criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage
};

await context.sync();
```

### [Text comparison](/javascript/api/excel/excel.textconditionalformat)

Text comparison conditional formatting uses string comparisons as the condition. The `rule` property is a [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule) defining a string to compare with the cell and an operator to specify the type of comparison. 

The following example formats the font color red when a cell's text contains "Delayed".

![A range with cells containing "Delayed" in red.](../images/excel-conditional-format-text.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B16:D18");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.containsText
);

// color the font of every cell containing "Delayed"
conditionalFormat.textComparison.format.font.color = "red";
conditionalFormat.textComparison.rule = {
     operator: Excel.ConditionalTextOperator.contains,
     text: "Delayed"
};

await context.sync();
```

### [Top/bottom](/javascript/api/excel/excel.TopBottomconditionalformat)

Top/bottom conditional formatting applies a format to the highest or lowest values in a range. The `rule` property, which is of type [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), sets whether the condition is based on the highest or lowest, as well as whether the evaluation is ranked or percentage-based. 

The following example applies a green highlight to the highest value cell in the range.


![A range with the highest number highlighted in green.](../images/excel-conditional-format-topbottom.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const range = sheet.getRange("B21:E23");
const conditionalFormat = range.conditionalFormats.add(
     Excel.ConditionalFormatType.topBottom
);

// for the highest valued cell in the range, make the background green
conditionalFormat.topBottom.format.fill.color = "green"
conditionalFormat.topBottom.rule = { rank: 1, type: "TopItems"}

await context.sync();
```

## Multiple formats and priority

You can apply multiple conditional formats to a range. If the formats have conflicting elements, such as differing font colors, only one format applies that particular element. Precedence is defined by the `ConditionalFormat.priority` property. Priority is a number (equal to the index in the `ConditionalFormatCollection`) and can be set when creating the format. The lowerer the `priority` value, the higher the priority of the format is.

The following example shows a conflicting font color choice between the two formats. Negative numbers will get a bold font, but NOT a red font, because priority goes to the format that gives them a blue font.

![A range with low numbers bolded and in red, negative numbers in blue with green backgrounds.](../images/excel-conditional-format-priority.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const temperatureDataRange = sheet.tables.getItem("TemperatureTable").getDataBodyRange();


// Set low numbers to bold, dark red font and assign priority 1.
const presetFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.presetCriteria);
presetFormat.preset.format.font.color = "red";
presetFormat.preset.format.font.bold = true;
presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };
presetFormat.priority = 1;

// Set negative numbers to blue font with green background and set priority 0.
const cellValueFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.cellValue);
cellValueFormat.cellValue.format.font.color = "blue";
cellValueFormat.cellValue.format.fill.color = "lightgreen";
cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
cellValueFormat.priority = 0;

await context.sync();

```

### Mutually exclusive conditional formats

The `stopIfTrue` property of `ConditionalFormat` prevents lower priority conditional formats from being applied to the range. When a range matching the conditional format with `stopIfTrue === true` is applied, no subsequent conditional formats are applied, even if their formatting details are not contradictory.

The following example shows two conditional formats being added to a range. Negative numbers will have a blue font with a light green background, regardless of whether the other format condition is true.

![A range with low numbers bolded and in red, unless they are negative, in which case they are not bolded, blue, and have a green background.](../images/excel-conditional-format-stopiftrue.png)

```typescript
const sheet = context.workbook.worksheets.getItem("Sample");
const temperatureDataRange = sheet.tables.getItem("TemperatureTable").getDataBodyRange();

// Set low numbers to bold, dark red font and assign priority 1.
const presetFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.presetCriteria);
presetFormat.preset.format.font.color = "red";
presetFormat.preset.format.font.bold = true;
presetFormat.preset.rule = { criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage };
presetFormat.priority = 1;

// Set negative numbers to blue font with green background and 
// set priority 0, but set stopIfTrue to true, so none of the 
// formatting of the conditional format with the higher priority
// value will apply, not even the bolding of the font.
const cellValueFormat = temperatureDataRange.conditionalFormats
    .add(Excel.ConditionalFormatType.cellValue);
cellValueFormat.cellValue.format.font.color = "blue";
cellValueFormat.cellValue.format.fill.color = "lightgreen";
cellValueFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
cellValueFormat.priority = 0;
cellValueFormat.stopIfTrue = true;

await context.sync();
```

## See also

- [Fundamental programming concepts with the Excel JavaScript API](../excel/excel-add-ins-core-concepts.md)
- [Work with ranges using the Excel JavaScript API](../excel/excel-add-ins-ranges.md)
- [ConditionalFormat Object (JavaScript API for Excel)](/javascript/api/excel/excel.conditionalformat)
- [Add, change, or clear conditional formats](https://support.office.com/article/add-change-or-clear-conditional-formats-8a1cc355-b113-41b7-a483-58460332a1af)
- [Use formulas with conditional formatting](https://support.office.com/article/Use-formulas-with-conditional-formatting-FED60DFA-1D3F-4E13-9ECB-F1951FF89D7F)
