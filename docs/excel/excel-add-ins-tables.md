---
title: Create, read, and manage Excel tables with JavaScript
description: Learn how to create, resize, read, sort, filter, and format Excel tables with the Excel JavaScript API.
ms.date: 06/03/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Create, read, and manage tables with the Excel JavaScript API

Use Excel tables when your add-in needs structured data that users can sort, filter, and format. This article shows how to create a table, add rows and columns, read table data, react to changes, and manage filters and formatting. For the full API surface, see [Table object](/javascript/api/excel/excel.table) and [TableCollection object](/javascript/api/excel/excel.tablecollection).

## Create a table

The following code sample creates a table in the worksheet named **Sample**. The table has headers and contains four columns and seven rows of data.

To name a table, create it first and then set its `name` property, as the following example shows.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
    ]);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();

    await context.sync();
});
```

### New table

:::image type="content" source="../images/excel-tables-create.png" alt-text="New table in Excel.":::

## Add rows to a table

The following code sample adds seven new rows to the table named **ExpensesTable** within the worksheet named **Sample**. The `index` parameter of the [`add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)) method is set to `null`, which specifies that the rows are added after the existing rows in the table. The `alwaysInsert` parameter is set to `true`, which indicates that the new rows are inserted into the table, not below the table. The code then sets the width of the columns and height of the rows to best fit the current data in the table.

The `index` property of a [TableRow](/javascript/api/excel/excel.tablerow) object indicates the row position in the table's `rows` collection. A `TableRow` object doesn't have an `id` property, so use the row position when you need to identify a row.

```js
// This code sample shows how to add rows to a table that already exists
// on a worksheet named Sample.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.rows.add(
        null, // index, Adds rows to the end of the table.
        [
            ["1/16/2017", "THE PHONE COMPANY", "Communications", "$120"],
            ["1/20/2017", "NORTHWIND ELECTRIC CARS", "Transportation", "$142"],
            ["1/20/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$27"],
            ["1/21/2017", "COHO VINEYARD", "Restaurant", "$33"],
            ["1/25/2017", "BELLOWS COLLEGE", "Education", "$350"],
            ["1/28/2017", "TREY RESEARCH", "Other", "$135"],
            ["1/31/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$97"]
        ],
        true, // alwaysInsert, Specifies that the new rows be inserted into the table.
    );

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

### Table with new rows

:::image type="content" source="../images/excel-tables-add-rows.png" alt-text="Table with new rows in Excel.":::

## Add a column to a table

These examples show how to add a column to a table. The first example populates the new column with static values. The second example populates the new column with formulas.

The `index` property of a [TableColumn](/javascript/api/excel/excel.tablecolumn) object indicates the column position in the table's `columns` collection. The `id` property of a `TableColumn` object contains a unique key that identifies the column.

### Add a column that contains static values

The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**. The code adds the new column after all existing columns in the table. The new column contains a header ("Day of the Week") and data to populate the cells in the column. The code sets the width of the columns and height of the rows to best fit the current data in the table.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Day of the Week"],
        ["Saturday"],
        ["Friday"],
        ["Monday"],
        ["Thursday"],
        ["Sunday"],
        ["Saturday"],
        ["Monday"]
    ]);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

#### Table with new column

:::image type="content" source="../images/excel-tables-add-column.png" alt-text="Table with new column in Excel.":::

### Add a column that contains formulas

The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**. The new column is added to the end of the table. It contains a header ("Type of the Day") and uses a formula to populate each data cell in the column. The width of the columns and height of the rows are then set to best fit the current data in the table.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Type of the Day"],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")']
    ]);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

#### Table with new calculated column

:::image type="content" source="../images/excel-tables-add-calculated-column.png" alt-text="Table with new calculated column in Excel.":::

## Resize a table

Your add-in can resize a table without adding data to the table or changing cell values. To resize a table, use the [Table.resize](/javascript/api/excel/excel.table#excel-excel-table-resize-member(1)) method. The following code sample shows how to resize a table. This code sample uses the **ExpensesTable** from the [Create a table](#create-a-table) section earlier in this article and sets the new range of the table to **A1:D20**.

```js
await Excel.run(async (context) => {
    // Retrieve the worksheet and a table on that worksheet.
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Resize the table.
    expensesTable.resize("A1:D20");

    await context.sync();
});
```

> [!IMPORTANT]
> The new range of the table must overlap with the original range, and the headers (or the top of the table) must be in the same row.

### Table after resize

:::image type="content" source="../images/excel-tables-resize.png" alt-text="Table with multiple empty rows in Excel.":::

## Rename a column

The following code sample renames the first column in the table to **Purchase date**. The code then sets the width of the columns and height of the rows to best fit the current data in the table.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    await context.sync();

    expensesTable.columns.items[0].name = "Purchase date";

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

### Table with new column name

:::image type="content" source="../images/excel-tables-update-column-name.png" alt-text="Table with new column name in Excel.":::

## Get data from a table

The following code sample reads data from a table named **ExpensesTable** in the worksheet named **Sample** and then outputs that data below the table in the same worksheet.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Get data from the header row.
    let headerRange = expensesTable.getHeaderRowRange().load("values");

    // Get data from the table.
    let bodyRange = expensesTable.getDataBodyRange().load("values");

    // Get data from a single column.
    let columnRange = expensesTable.columns.getItem("Merchant").getDataBodyRange().load("values");

    // Get data from a single row.
    let rowRange = expensesTable.rows.getItemAt(1).load("values");

    // Sync to populate proxy objects with data from Excel.
    await context.sync();

    let headerValues = headerRange.values;
    let bodyValues = bodyRange.values;
    let merchantColumnValues = columnRange.values;
    let secondRowValues = rowRange.values;

    // Write data from table back to the sheet
    sheet.getRange("A11:A11").values = [["Results"]];
    sheet.getRange("A13:D13").values = headerValues;
    sheet.getRange("A14:D20").values = bodyValues;
    sheet.getRange("B23:B29").values = merchantColumnValues;
    sheet.getRange("A32:D32").values = secondRowValues;

    // Sync to update the sheet in Excel.
    await context.sync();
});
```

### Table and data output

:::image type="content" source="../images/excel-tables-get-data.png" alt-text="Table data in Excel.":::

## Detect data changes

Your add-in might need to react when users change the data in a table. To detect these changes, [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a table. Event handlers for the `onChanged` event receive a [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) object when the event fires.

The `TableChangedEventArgs` object provides information about the changes and the source. Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values actually changed. The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail). The following code sample shows how to display the before and after values and types of a cell that changed.

```js
// This function would be used as an event handler for the Table.onChanged event.
async function onTableChanged(eventArgs) {
    await Excel.run(async (context) => {
        let details = eventArgs.details;
        let address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        await context.sync();
    });
}
```

## Sort data in a table

The following code sample sorts table data in descending order according to the values in the fourth column of the table.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to sort data by the fourth column of the table (descending).
    let sortRange = expensesTable.getDataBodyRange();
    sortRange.sort.apply([
        {
            key: 3,
            ascending: false,
        },
    ]);

    // Sync to run the queued command in Excel.
    await context.sync();
});
```

### Table data sorted by Amount (descending)

:::image type="content" source="../images/excel-tables-sort.png" alt-text="Sorted table data in Excel.":::

When you sort data in a worksheet, an event notification fires. To learn more about sort-related events and how your add-in can register event handlers to respond to such events, see [Handle sorting events](excel-add-ins-worksheets.md#handle-sorting-events).

## Apply filters to a table

The following code sample applies filters to the **Amount** column and the **Category** column within a table. As a result of the filters, only rows where **Category** is one of the specified values and **Amount** is below the average value for all rows is shown.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to apply a filter on the Category column.
    let categoryFilter = expensesTable.columns.getItem("Category").filter;
    categoryFilter.apply({
      filterOn: Excel.FilterOn.values,
      values: ["Restaurant", "Groceries"]
    });

    // Queue a command to apply a filter on the Amount column.
    let amountFilter = expensesTable.columns.getItem("Amount").filter;
    amountFilter.apply({
      filterOn: Excel.FilterOn.dynamic,
      dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    // Sync to run the queued commands in Excel.
    await context.sync();
});
```

### Table data with filters applied for Category and Amount

:::image type="content" source="../images/excel-tables-filters-apply.png" alt-text="Table data filtered in Excel.":::

## Clear table filters

The following code sample clears any filters currently applied on the table.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    await context.sync();
});
```

### Table data with no filters applied

:::image type="content" source="../images/excel-tables-filters-clear.png" alt-text="Unfiltered table data in Excel.":::

## Get visible cells from a filtered table

The following code sample gets only the cells that are currently visible in the specified table and then writes those values to the console. Use the `getVisibleView()` method when your add-in needs the visible contents of a filtered table.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    let visibleRange = expensesTable.getDataBodyRange().getVisibleView();
    visibleRange.load("values");

    await context.sync();
    console.log(visibleRange.values);
});
```

## Filter a table with AutoFilter

Your add-in can use the table's [AutoFilter](/javascript/api/excel/excel.autofilter) object to filter data. An `AutoFilter` object represents the full filter state of a table or range. Because it gives you a single access point for filters, it can be easier to manage multiple filters together.

The following code sample shows the same [data filtering as the earlier code sample](#apply-filters-to-a-table), but done entirely through the auto-filter.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.autoFilter.apply(expensesTable.getRange(), 2, {
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });
    expensesTable.autoFilter.apply(expensesTable.getRange(), 3, {
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    await context.sync();
});
```

You can also apply an `AutoFilter` to a range at the worksheet level. For more information, see [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#filter-data).

## Format a table

The following code sample applies formatting to a table. It specifies different fill colors for the header row of the table, the body of the table, the second row of the table, and the first column of the table. For information about the properties you can use to specify format, see [RangeFormat Object (JavaScript API for Excel)](/javascript/api/excel/excel.rangeformat).

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.getHeaderRowRange().format.fill.color = "#C70039";
    expensesTable.getDataBodyRange().format.fill.color = "#DAF7A6";
    expensesTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";
    expensesTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";

    await context.sync();
});
```

### Table after formatting is applied

:::image type="content" source="../images/excel-tables-formatting-after.png" alt-text="Table after formatting is applied in Excel.":::

## Convert a range to a table

The following code sample creates a range of data and then converts that range to a table. The width of the columns and height of the rows are then set to best fit the current data in the table.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Define values for the range.
    let values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
    ["Frames", 5000, 7000, 6544, 4377],
    ["Saddles", 400, 323, 276, 651],
    ["Brake levers", 12000, 8766, 8456, 9812],
    ["Chains", 1550, 1088, 692, 853],
    ["Mirrors", 225, 600, 923, 544],
    ["Spokes", 6005, 7634, 4589, 8765]];

    // Create the range.
    let range = sheet.getRange("A1:E7");
    range.values = values;

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();

    // Convert the range to a table.
    let expensesTable = sheet.tables.add('A1:E7', true);
    expensesTable.name = "ExpensesTable";

    await context.sync();
});
```

### Data in the range (before the range is converted to a table)

:::image type="content" source="../images/excel-ranges.png" alt-text="Data in range in Excel.":::

### Data in the table (after the range is converted to a table)

:::image type="content" source="../images/excel-tables-from-range.png" alt-text="Data in table in Excel.":::

## Import JSON data into a table

The following code sample creates a table in the worksheet named **Sample** and then populates the table by using a JSON object that defines two rows of data. The width of the columns and height of the rows are then set to best fit the current data in the table.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    let transactions = [
      {
        "DATE": "1/1/2017",
        "MERCHANT": "The Phone Company",
        "CATEGORY": "Communications",
        "AMOUNT": "$120"
      },
      {
        "DATE": "1/1/2017",
        "MERCHANT": "Southridge Video",
        "CATEGORY": "Entertainment",
        "AMOUNT": "$40"
      }
    ];

    let newData = transactions.map(item =>
        [item.DATE, item.MERCHANT, item.CATEGORY, item.AMOUNT]);

    expensesTable.rows.add(null, newData);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();

    await context.sync();
});
```

### New table

:::image type="content" source="../images/excel-tables-create-from-json.png" alt-text="New table from imported JSON data in Excel.":::

## See also

- [Core Excel object model concepts for Office Add-ins](excel-add-ins-core-concepts.md)
- [Get Excel worksheet ranges with the JavaScript API](excel-add-ins-ranges-get.md)
- [Add data validation to Excel ranges](excel-add-ins-data-validation.md)
- [Create and customize charts with the Excel JavaScript API](excel-add-ins-charts.md)
- [Performance optimization using the Excel JavaScript API](performance.md)
