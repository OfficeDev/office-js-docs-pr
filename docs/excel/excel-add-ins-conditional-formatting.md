---
title: Apply conditional formatting to ranges with the Excel JavaScript API
description: Learn about conditional formatting in the context of Excel JavaScript add-ins.
ms.date: 05/19/2023
ms.localizationpriority: medium
---

# Apply conditional formatting to Excel ranges

The Excel JavaScript Library provides APIs to apply conditional formatting to data ranges in your worksheets. This functionality makes large sets of data easy to visually parse. The formatting also dynamically updates based on changes within the range.

> [!NOTE]
> This article covers conditional formatting in the context of Excel JavaScript add-ins. The following articles provide detailed information about the full conditional formatting capabilities within Excel.
>
> - [Add, change, or clear conditional formats](https://support.microsoft.com/office/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
> - [Use formulas with conditional formatting](https://support.microsoft.com/office/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)

## Programmatic control of conditional formatting

The `Range.conditionalFormats` property is a collection of [ConditionalFormat](/javascript/api/excel/excel.conditionalformat) objects that apply to the range.  The `ConditionalFormat` object contains several properties that define the format to be applied based on the [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype).

- `cellValue`
- `colorScale`
- `custom`
- `dataBar`
- `iconSet`
- `preset`
- `textComparison`
- `topBottom`

> [!NOTE]
> Each of these formatting properties has a corresponding `*OrNullObject` variant. Learn more about that pattern in the [\*OrNullObject methods](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) section.

Only one format type can be set for the ConditionalFormat object. This is determined by the `type` property, which is a [ConditionalFormatType](/javascript/api/excel/excel.conditionalformattype) enum value. `type` is set when adding a conditional format to a range.

## Create conditional formatting rules

Conditional formats are added to a range by using `conditionalFormats.add`. Once added, the properties specific to the conditional format can be set. The following examples show the creation of different formatting types.

### [Cell value](/javascript/api/excel/excel.cellvalueconditionalformat)

Cell value conditional formatting applies a user-defined format based on the results of one or two formulas in the [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule). The `operator` property is a [ConditionalCellValueOperator](/javascript/api/excel/excel.conditionalcellvalueoperator) defining how the resulting expressions relate to the formatting.

The following example shows red font coloring applied to any value in the range less than zero.

:::image type="content" source="../images/excel-conditional-format-cell-value.png" alt-text="A range with negative numbers in red.":::

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B21:E23");
    const conditionalFormat = range.conditionalFormats.add(
        Excel.ConditionalFormatType.cellValue
    );
    
    // Set the font of negative numbers to red.
    conditionalFormat.cellValue.format.font.color = "red";
    conditionalFormat.cellValue.rule = { formula1: "=0", operator: "LessThan" };
    
    await context.sync();
});
```

### [Color scale](/javascript/api/excel/excel.colorscaleconditionalformat)

Color scale conditional formatting applies a color gradient across the data range. The `criteria` property on the `ColorScaleConditionalFormat` defines three [ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion): `minimum`, `maximum`, and, optionally, `midpoint`. Each of the criterion scale points have three properties:

- `color` - The HTML color code for the endpoint.
- `formula` - A number or formula representing the endpoint. This will be `null` if `type` is `lowestValue` or `highestValue`.
- `type` - How the formula should be evaluated. `highestValue` and `lowestValue` refer to values in the range being formatted.

The following example shows a range being colored blue to yellow to red. Note that `minimum` and `maximum` are the lowest and highest values respectively and use `null` formulas. `midpoint` is using the `percentage` type with a formula of `"=50"` so the yellowest cell is the mean value.

:::image type="content" source="../images/excel-conditional-format-color-scale.png" alt-text="A range with the low number in blue, average number in yellow, and high number is red, with gradients for between values.":::

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:M5");
    const conditionalFormat = range.conditionalFormats.add(
          Excel.ConditionalFormatType.colorScale
    );
    
    // Color the backgrounds of the cells from blue to yellow to red based on value.
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
});
```

### [Custom](/javascript/api/excel/excel.customconditionalformat)

Custom conditional formatting applies a user-defined format to the cells based on a formula of arbitrary complexity. The [ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule) object lets you define the formula in different notations:

- `formula` - Standard notation.
- `formulaLocal` - Localized based on the user's language.
- `formulaR1C1` - R1C1-style notation.

The following example colors the fonts green of cells with higher values than the cell to their left.

:::image type="content" source="../images/excel-conditional-format-custom.png" alt-text="A range with green numbers for places the preceding column's value in that row is lower.":::

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B8:E13");
    const conditionalFormat = range.conditionalFormats.add(
         Excel.ConditionalFormatType.custom
    );
    
    // If a cell has a higher value than the one to its left, set that cell's font to green.
    conditionalFormat.custom.rule.formula = '=IF(B8>INDIRECT("RC[-1]",0),TRUE)';
    conditionalFormat.custom.format.font.color = "green";
    
    await context.sync();
});

```

### [Data bar](/javascript/api/excel/excel.databarconditionalformat)

Data bar conditional formatting adds data bars to the cells. By default, the minimum and maximum values in the Range form the bounds and proportional sizes of the data bars. The `DataBarConditionalFormat` object has several properties to control the bar's appearance.

The following example formats the range with data bars filling left-to-right.

:::image type="content" source="../images/excel-conditional-format-databar.png" alt-text="A range with databars behind the values in cells.":::

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B8:E13");
    const conditionalFormat = range.conditionalFormats.add(
         Excel.ConditionalFormatType.dataBar
    );
    
    // Give left-to-right, default-appearance data bars to all the cells.
    conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
    await context.sync();
});
```

### [Icon set](/javascript/api/excel/excel.iconsetconditionalformat)

Icon set conditional formatting uses Excel [Icons](/javascript/api/excel/excel.icon) to highlight cells. The `criteria` property is an array of [ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion), which define the symbol to be inserted and the condition under which it is inserted. This array is automatically prepopulated with criterion elements with default properties. Individual properties cannot be overwritten. Instead, the whole criteria object must be replaced.

The following example shows a three-triangle icon set applied across the range.

:::image type="content" source="../images/excel-conditional-format-iconset.png" alt-text="A range with green upward triangles for values above 1000, yellow lines for values between 700 and 1000, and red downward triangles for lower values.":::

```js
await Excel.run(async (context) => {
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
        {},
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
});
```

### [Preset criteria](/javascript/api/excel/excel.presetcriteriaconditionalformat)

Preset conditional formatting applies a user-defined format to the range based on a selected standard rule. These rules are defined by the [ConditionalFormatPresetCriterion](/javascript/api/excel/excel.conditionalformatpresetcriterion) in the [ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule).

The following example colors the font white wherever a cell's value is at least one standard deviation above the range's average.

:::image type="content" source="../images/excel-conditional-format-preset.png" alt-text="A range with white font cells where the values are at least one standard deviation above average.":::

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:M5");
    const conditionalFormat = range.conditionalFormats.add(
         Excel.ConditionalFormatType.presetCriteria
    );
    
    // Color every cell's font white that is one standard deviation above average relative to the range.
    conditionalFormat.preset.format.font.color = "white";
    conditionalFormat.preset.rule = {
         criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage
    };
    
    await context.sync();
});
```

### [Text comparison](/javascript/api/excel/excel.textconditionalformat)

Text comparison conditional formatting uses string comparisons as the condition. The `rule` property is a [ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule) defining a string to compare with the cell and an operator to specify the type of comparison.

The following example formats the font color red when a cell's text contains "Delayed".

:::image type="content" source="../images/excel-conditional-format-text.png" alt-text="A range with cells containing 'Delayed' in red.":::

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B16:D18");
    const conditionalFormat = range.conditionalFormats.add(
         Excel.ConditionalFormatType.containsText
    );
    
    // Color the font of every cell containing "Delayed".
    conditionalFormat.textComparison.format.font.color = "red";
    conditionalFormat.textComparison.rule = {
         operator: Excel.ConditionalTextOperator.contains,
         text: "Delayed"
    };
    
    await context.sync();
});
```

### [Top/bottom](/javascript/api/excel/excel.topbottomconditionalformat)

Top/bottom conditional formatting applies a format to the highest or lowest values in a range. The `rule` property, which is of type [ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule), sets whether the condition is based on the highest or lowest, as well as whether the evaluation is ranked or percentage-based.

The following example applies a green highlight to the highest value cell in the range.

:::image type="content" source="../images/excel-conditional-format-topbottom.png" alt-text="A range with the highest number highlighted in green.":::

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B21:E23");
    const conditionalFormat = range.conditionalFormats.add(
         Excel.ConditionalFormatType.topBottom
    );
    
    // For the highest valued cell in the range, make the background green.
    conditionalFormat.topBottom.format.fill.color = "green"
    conditionalFormat.topBottom.rule = { rank: 1, type: "TopItems"}
    
    await context.sync();
});
```

## Change conditional formatting rules

The `ConditionalFormat` object offers multiple methods to change conditional formatting rules after they've been set.

- [changeRuleToCellValue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocellvalue-member(1))
- [changeRuleToColorScale](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocolorscale-member(1))
- [changeRuleToContainsText](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocontainstext-member(1))
- [changeRuleToCustom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocustom-member(1))
- [changeRuleToDataBar](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletodatabar-member(1))
- [changeRuleToIconSet](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletoiconset-member(1))
- [changeRuleToPresetCriteria](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletopresetcriteria-member(1))
- [changeRuleToTopBottom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletotopbottom-member(1))

The following example shows how to use the `changeRuleToPresetCriteria` method from the preceding list to change an existing conditional format rule to the preset criteria rule type.

> [!NOTE]
> The specified range must have an existing conditional format rule to use the change methods. If the specified range has no conditional format rule, the change methods don't apply a new rule.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:M5");
    
    // Retrieve the first existing `ConditionalFormat` rule on this range. 
    // Note: The specified range must have an existing conditional format rule.
    const conditionalFormat = range.conditionalFormats.getItemOrNullObject("0");
    
    // Change the conditional format rule to preset criteria.
    conditionalFormat.changeRuleToPresetCriteria({
        criterion: Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage, 
    });
    conditionalFormat.preset.format.font.color = "red";
    
    await context.sync();
});
```

## Multiple formats and priority

You can apply multiple conditional formats to a range. If the formats have conflicting elements, such as differing font colors, only one format applies that particular element. Precedence is defined by the `ConditionalFormat.priority` property. Priority is a number (equal to the index in the `ConditionalFormatCollection`) and can be set when creating the format. The lower the `priority` value, the higher the priority of the format is.

The following example shows a conflicting font color choice between the two formats. Negative numbers will get a bold font, but NOT a red font, because priority goes to the format that gives them a blue font.

:::image type="content" source="../images/excel-conditional-format-priority.png" alt-text="A range with low numbers bolded and in red, negative numbers in blue with green backgrounds.":::

```js
await Excel.run(async (context) => {
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
});
```

### Mutually exclusive conditional formats

The `stopIfTrue` property of `ConditionalFormat` prevents lower priority conditional formats from being applied to the range. When a range matching the conditional format with `stopIfTrue === true` is applied, no subsequent conditional formats are applied, even if their formatting details are not contradictory.

The following example shows two conditional formats being added to a range. Negative numbers will have a blue font with a light green background, regardless of whether the other format condition is true.

:::image type="content" source="../images/excel-conditional-format-stopiftrue.png" alt-text="A range with low numbers bolded and in red, unless they are negative, in which case they are not bolded, blue, and have a green background.":::

```js
await Excel.run(async (context) => {
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
});
```

## Clear conditional formatting rules

To remove format properties from a specific conditional format rule, use the [clearFormat](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-clearformat-member(1)) method of the `ConditionalRangeFormat` object. The `clearFormat` method creates a formatting rule without format settings.

To remove all the conditional formatting rules from a specific range, or an entire worksheet, use the [clearAll](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-clearall-member(1)) method of the `ConditionalFormatCollection` object.

The following sample shows how to remove all conditional formatting from a worksheet with the `clearAll` method.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange();
    range.conditionalFormats.clearAll();

    await context.sync();
});
```

## See also

- [Excel JavaScript object model in Office Add-ins](../excel/excel-add-ins-core-concepts.md)
- [ConditionalFormat Object (JavaScript API for Excel)](/javascript/api/excel/excel.conditionalformat)
- [Add, change, or clear conditional formats](https://support.microsoft.com/office/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
- [Use formulas with conditional formatting](https://support.microsoft.com/office/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
