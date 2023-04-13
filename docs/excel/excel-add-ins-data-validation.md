---
title: Add data validation to Excel ranges
description: Learn how the Excel JavaScript APIs enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.
ms.date: 02/16/2022
ms.localizationpriority: medium
---

# Add data validation to Excel ranges

The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook. To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI.

- [Apply data validation to cells](https://support.microsoft.com/office/29fecbcc-d1b9-42c1-9d76-eff3ce5f7249)
- [More on data validation](https://support.microsoft.com/office/f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Description and examples of data validation in Excel](https://support.microsoft.com/help/211485)

## Programmatic control of data validation

The `Range.dataValidation` property, which takes a [DataValidation](/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel. There are five properties to the `DataValidation` object:

- `rule` &#8212; Defines what constitutes valid data for the range. See [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, `information`, `warning`, and `stop`. See [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message. See [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range. Defaults to `true`.
- `type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.

> [!NOTE]
> Data validation added programmatically behaves just like manually added data validation. In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option. If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.

## Creating validation rules

To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`. This takes a [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties. *No more than one of these properties may be present in any `DataValidationRule` object.* The property that you include determines the type of validation.

### Basic and DateTime validation rule types

The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) object as their value.

- `wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.
- `decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.
- `textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.

Here is an example of creating a validation rule. Note the following about this code.

- The `operator` is the binary operator `greaterThan`. Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand. So this rule says that only whole numbers that are greater than 0 are valid.
- The `formula1` is a hard-coded number. If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value. For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: Excel.DataValidationOperator.greaterThan
            }
        };

    await context.sync();
});
```

See [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.

There are also two ternary operators: `between` and `notBetween`. To use these, you must specify the optional `formula2` property. The `formula1` and `formula2` values are the bounding operands. The value that the user tries to enter in the cell is the third (evaluated) operand. The following is an example of using the "Between" operator.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
              operator: Excel.DataValidationOperator.between
            }
        };

    await context.sync();
});
```

The next two rule properties take a [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) object as their value.

- `date`
- `time`

The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way. The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula). The following is an example that defines valid values as dates in the first week of April, 2022.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            date: {
                formula1: "2022-04-01",
                formula2: "2022-04-08",
                operator: Excel.DataValidationOperator.between
            }
        };

    await context.sync();
});
```

### List validation rule type

Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list. The following is an example. Note the following about this code.

- It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.
- The `source` property specifies the list of valid values. The string argument refers to a range containing the names. You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".
- The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it. If set to `true`, then the drop-down appears with the list of values from the `source`.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");   
    let nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: "=Names!$A$1:$A$3"
        }
    };

    await context.sync();
})
```

### Custom validation rule type

Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula. The following is an example. Note the following about this code.

- It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.
- To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.
- `SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2. If A2 is not contained in B2, it does not return a number. `ISNUMBER()` returns a boolean. So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let commentsRange = sheet.tables.getItem("AthletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    await context.sync();
});
```

## Create validation error alerts

You can a create custom error alert that appears when a user tries to enter invalid data in a cell. The following is a simple example. Note the following about this code.

- The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert. Only `stop` actually prevents the user from adding invalid data. The pop-ups for `warning` and `information` have options that allow the user enter the invalid data anyway.
- The `showAlert` property defaults to `true`. This means that Excel will pop-up a generic alert (of type `stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style. This code sets a custom message and title.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // The default is 'true'.
              style: Excel.DataValidationAlertStyle.stop,
            title: "Negative or Decimal Number Entered"
        };

    // Set range.dataValidation.rule and optionally .prompt here.

    await context.sync();
});
```

For more information, see [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).

## Create validation prompts

You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied. The following is an example.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // The default is 'false'.
            title: "Positive Whole Numbers Only."
        };

    // Set range.dataValidation.rule and optionally .errorAlert here.

    await context.sync();
});
```

For more information, see [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).

## Remove data validation from a range

To remove data validation from a range, call the  [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-clear-member(1)) method.

```js
myrange.dataValidation.clear()
```

It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation. If it isn't, only the overlapping cells, if any, of the two ranges are cleared. 

> [!NOTE]
> Clearing data validation from a range will also clear any data validation that a user has added manually to the range.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [DataValidation Object (JavaScript API for Excel)](/javascript/api/excel/excel.datavalidation)
- [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range)
