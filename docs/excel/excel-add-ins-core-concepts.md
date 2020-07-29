---
title: Fundamental programming concepts with the Excel JavaScript API
description: 'Use the Excel JavaScript API to build add-ins for Excel.'
ms.date: 07/28/2020
localization_priority: Priority
---

# Fundamental programming concepts with the Excel JavaScript API

This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later. It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.

> [!IMPORTANT]
> Please read [Using the host-specific API model](../develop/host-specific-api-model.md) to learn about the asynchronous nature of the Excel APIs and how they work with the workbook.  

## Office.js APIs for Excel

An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:

* **Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.

* **Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.

While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API. For example:

* [Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.
* [Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.

The following image illustrates when you might use the Excel JavaScript API or the Common APIs.

![Image of the differences between the Excel JS API and Common APIs](../images/excel-js-api-common-api.png)

## Object model

To understand the Excel APIs, you must understand how the components of a workbook are related to one another.

* A **Workbook** contains one or more **Worksheets**.
* A **Worksheet** gives access to cells through **Range** objects.
* A **Range** represents a group of contiguous cells.
* **Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.
* A **Worksheet** contains collections of those data objects that are present in the individual sheet.
* **Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.

### Ranges

A range is a group of contiguous cells in the workbook. Add-ins typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.

Ranges have three core properties: `values`, `formulas`, and `format`. These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.

#### Range sample

The following sample shows how to create sales records. This function uses `Range` objects to set the values, formulas, and formats.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

This sample creates the following data in the current worksheet:

![A sales record showing value rows, a formula column, and formatted headers.](../images/excel-overview-range-sample.png)

### Charts, tables, and other data objects

The Excel JavaScript APIs can create and manipulate the data structures and visualizations within Excel. Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.

#### Creating a table

Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.

The following sample creates a table using the ranges from the previous sample.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

Using this sample code on the worksheet with the previous data creates the following table:

![A table made from the previous sales record.](../images/excel-overview-table-sample.png)

#### Creating a chart

Create charts to visualize the data in a range. The APIs support dozens of chart varieties, each of which can be customized to suit your needs.

The following sample creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

Running this sample on the worksheet with the previous table creates the following chart:

![A column chart showing quantities of three items from the previous sales record.](../images/excel-overview-chart-sample.png)

## Run options

`Excel.run` has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object. This contains a set of properties that affect platform behavior when the function runs. The following property is currently supported:

* `delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode. When **true**, the batch request is delayed and runs when the user exits cell edit mode. When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user). The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## null or blank property values

`null` and empty strings have special implications in the Excel JavaScript APIs. There are used to represent empty cells, no formatting, or default values. This section details the use of `null` and empty string when getting and setting properties.

### null input in 2-D Array

In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.

For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### null input for a property

`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.

```js
range.values = null;
```

Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.

```js
range.format.fill.color =  null;
```

### null property values in the response

Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:

* If all cells in the range have the same font color, `range.format.font.color` specifies that color.
* If multiple font colors are present within the range, `range.format.font.color` is `null`.

### Blank input for a property

When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:

* If you specify a blank value for the `values` property of a range, the content of the range is cleared.
* If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.
* If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.

### Blank property values in the response

For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## Requirement sets

Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs. To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).

### Checking for requirement set support at runtime

The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### Defining requirement set support in the manifest

You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.

The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> To make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.

### Requirement sets for the Office.js Common API

For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).

## Handle errors

When an API error occurs, the API returns an `error` object that contains a code and a message. For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).

## See also

* [Build your first Excel add-in](../quickstarts/excel-quickstart-jquery.md)
* [Excel add-ins code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API performance optimization](../excel/performance.md)
* [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
* [Common coding issues and unexpected platform behaviors](../develop/common-coding-issues.md)
