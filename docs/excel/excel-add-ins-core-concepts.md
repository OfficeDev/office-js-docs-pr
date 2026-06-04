---
title: Core Excel object model concepts
description: Learn how workbooks, worksheets, ranges, tables, and charts relate in the Excel JavaScript API so you can read, write, and visualize workbook data.
ms.date: 06/03/2026
ms.topic: article
ms.localizationpriority: high
ai-usage: ai-assisted
---

# Core Excel object model concepts for Office Add-ins

This article explains how workbooks, worksheets, ranges, tables, and charts fit together in the Excel JavaScript object model.

In most Excel add-ins, you start with a workbook, move to a worksheet, work with one or more ranges, and then create higher-level objects such as tables or charts. Understanding that flow helps you develop your add-in faster and choose the right API for each task.

> [!IMPORTANT]
> Before you start with Excel-specific APIs, learn how `Excel.run`, proxy objects, and `context.sync()` work in [Application-specific API model](../develop/application-specific-api-model.md).

## Start with the Excel objects you'll use most

Excel add-ins typically start with the workbook and work to more specific elements of the spreadsheet. Here's how to conceptualize some of the JavaScript objects.

- A `Workbook` contains one or more `Worksheet` objects.
- A `Worksheet` contains cells and sheet-level objects.
- A `Range` represents one cell or a block of contiguous cells.
- `Range` objects are the starting point for writing values, formulas, and formats.
- `Table` and `Chart` objects are usually created from data that already exists in a range.

If you're new to the Excel object model, start with these common tasks:

- [Set and get range values, text, or formulas](excel-add-ins-ranges-set-get-values.md)
- [Work with tables](excel-add-ins-tables.md)
- [Work with worksheets](excel-add-ins-worksheets.md)

## Work with ranges

A range is a group of contiguous cells in a workbook. Add-ins usually use A1-style notation to define ranges, such as **B3** for a single cell or **C2:F4** for a rectangular block of cells.

Ranges expose three core properties that most add-ins use right away:

- `values` to read or write cell values
- `formulas` to read or write formulas
- `format` to change visual formatting

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

### Build a simple sales worksheet with ranges

The following example creates a small sales report. It writes a header row and product rows, calculates totals with formulas, and formats the totals as currency. Use this pattern when your add-in needs to populate and format a block of cells in one operation.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    const headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    const headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    const productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.56]
    ];
    const dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    const totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];
    const totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    await context.sync();
});
```

This sample creates the following data in the active worksheet.

:::image type="content" source="../images/excel-overview-range-sample.png" alt-text="A sales record showing value rows, a formula column, and formatted headers.":::

For more information, see [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md).

## Turn ranges into tables and charts

After your add-in writes data to a range, it often turns that data into a richer object. Tables make data easier to sort and filter. Charts make patterns easier to understand at a glance.

The Excel JavaScript API also supports other workbook objects, including PivotTables, shapes, and images. However, tables and charts are the most common next step after you create a range.

### Create a table from a range

Create a table when users need built-in filtering, structured references, and table formatting. The following example converts the sales data from the previous sample into a table.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    await context.sync();
});
```

When you run this code on the worksheet with the previous data, Excel creates the following table.

:::image type="content" source="../images/excel-overview-table-sample.png" alt-text="A table made from the previous sales record.":::

For more information, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).

### Create a chart from a range

Create a chart when you want users to interpret workbook data visually. The following example creates a stacked column chart from the item and quantity data, then places the chart 100 pixels below the top of the worksheet.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    await context.sync();
});
```

When you run this code on the worksheet with the previous table, Excel creates the following chart.

:::image type="content" source="../images/excel-overview-chart-sample.png" alt-text="A column chart showing quantities of three items from the previous sales record.":::

For more information, see [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).

## Know when to use Common APIs

[!include[The roles of the Common and application-specific APIs](../includes/excel-api-models.md)]

You'll use the Excel JavaScript API for most workbook operations, but you'll also use objects in the Common API for add-in runtime information and file access.

- [Context](/javascript/api/office/office.context): Use the `Context` object to inspect the add-in runtime, including `contentLanguage`, `officeTheme`, `host`, and `platform`. You can also call `requirements.isSetSupported()` to check whether Excel supports a specific requirement set.
- [Document](/javascript/api/office/office.document): Use the `Document` object and its `getFileAsync()` method when you need to download the workbook file where the add-in is running.

The following image shows when you might use the Excel JavaScript API instead of the Common APIs.

:::image type="content" source="../images/excel-js-api-common-api.png" alt-text="Differences between the Excel JS API and Common APIs.":::

## See also

- [Overview of Excel add-ins](excel-add-ins-overview.md)
- [Build your first Excel add-in](../quickstarts/excel-quickstart-jquery.md)
- [Work with worksheets](excel-add-ins-worksheets.md)
- [Set and get range values, text, or formulas](excel-add-ins-ranges-set-get-values.md)
- [Work with tables](excel-add-ins-tables.md)
- [Excel add-ins code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Optimize Excel JavaScript API performance](performance.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
