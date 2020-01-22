---
title: Work with PivotTables using the Excel JavaScript API
description: Use the Excel JavaScript API to create PivotTables and interact with their components. 
ms.date: 01/22/2020
localization_priority: Normal
---

# Work with PivotTables using the Excel JavaScript API

PivotTables streamline larger data sets. They allow the quick manipulation of grouped data. The Excel JavaScript API lets your add-in create PivotTables and interact with their components. This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.

If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.
See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.

> [!IMPORTANT]
> PivotTables created with OLAP are not currently supported. There is also no support for Power Pivot.

## Object model

The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.

- `Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.
- A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).
- A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield). If the design expands to include OLAP PivotTables, this may change.
- A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).
- A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.

Let's look at how these relationships apply to some example data. The following data describes fruit sales from various farms. It will be the example throughout this article.

![A collection of fruit sales of different types from different farms.](../images/excel-pivots-raw-data.png)

This fruit farm sales data will be used to make a PivotTable. Each column, such as **Types**, is a `PivotHierarchy`. The **Types** hierarchy contains the **Types** field. The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.

### Hierarchies

PivotTables are organized based on four hierarchy categories: row, column, data, and filter.

The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**. Each hierarchy can only exist in one of the four categories. If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.

Row and column hierarchies define how data will be grouped. For example, a row hierarchy of **Farms** will group together all the data sets from the same farm. The choice between row and column hierarchy defines the orientation of the PivotTable.

Data hierarchies are the values to be aggregated based on the row and column hierarchies. A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.

Filter hierarchies include or exclude data from the pivot based on values within that filtered type. A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.

Here is the farm data again, alongside a PivotTable. The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).

![A selection of fruit sales data next to a PivotTable with row, data, and filter hierarchies.](../images/excel-pivot-table-and-data.png)

This PivotTable could be generated through the JavaScript API or through the Excel UI. Both options allow for further manipulation through add-ins.

## Create a PivotTable

PivotTables need a name, source, and destination. The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type). The destination is a range address (given as either a `Range` or `string`).
The following samples show various PivotTable creation techniques.

### Create a PivotTable with range addresses

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### Create a PivotTable with Range objects

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21.
    var rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    var rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
      "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    return context.sync();
});
```

### Create a PivotTable at the workbook level

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## Use an existing PivotTable

Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets. The following code gets a PivotTable named  **My Pivot** from the workbook.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## Add rows and columns to a PivotTable

Rows and columns pivot the data around those fields’ values.

Adding the **Farm** column pivots all the sales around each farm. Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.

![A PivotTable with a Farm column and Type and Classification rows.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

You can also have a PivotTable with only rows or columns.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## Add data hierarchies to the PivotTable

Data hierarchies fill the PivotTable with information to combine based on the rows and columns. Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.

In the example, both **Farm** and **Type** are rows, with the crate sales as the data.

![A PivotTable showing the total sales of different fruit based on the farm they came from.](../images/excel-pivots-data-hierarchy.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based.
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case).
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    return context.sync();
});
```

## Slicers

[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table. A slicer uses values from a specified column or PivotField to filter corresponding rows. These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`. Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)). The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.

![A slicer filtering data on a PivotTable.](../images/excel-slicer.png)

> [!NOTE]
> The techniques described in this section focus on how to use slicers connected to PivotTables. The same techniques also apply to using slicers connected to tables.

### Create a slicer

You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method. Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object. The `SlicerCollection.add` method has three parameters:

- `slicerSource`: The data source on which the new slicer is based. It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.
- `sourceField`: The field in the data source by which to filter. It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.
- `slicerDestination`: The worksheet where the new slicer will be created. It can be a `Worksheet` object or the name or ID of a `Worksheet`. This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`. In this case, the collection's worksheet is used as the destination.

The following code sample adds a new slicer to the **Pivot** worksheet. The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data. The slicer is also named **Fruit Slicer** for future reference.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Pivot");
    var slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    return context.sync();
});
```

### Filter items with a slicer

The slicer filters the PivotTable with items from the `sourceField`. The `Slicer.selectItems` method sets the items that remain in the slicer. These items are passed to the method as a `string[]`, representing the keys of the items. Any rows containing those items remain in the PivotTable's aggregation. Subsequent calls to `selectItems` set the list to the keys specified in those calls.

> [!NOTE]
> If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown. The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).

The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

### Style and format a slicer

You add-in can adjust a slicer's display settings through `Slicer` properties. The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.caption = "Fruit Types";
    slicer.left = 395;
    slicer.top = 15;
    slicer.height = 135;
    slicer.width = 150;
    slicer.style = "SlicerStyleLight6";
    return context.sync();
});
```

### Delete a slicer

To delete a slicer, call the `Slicer.delete` method. The following code sample deletes the first slicer from the current worksheet.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## Change aggregation function

Data hierarchies have their values aggregated. For datasets of numbers, this is a sum by default. The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.

The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).

The following code samples changes the aggregation to be averages of the data.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    return context.sync().then(function() {

        // Change the aggregation from the default sum to an average of all the values in the hierarchy.
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        return context.sync();
    });
});
```

## Change calculations with a ShowAsRule

PivotTables, by default, aggregate the data of their row and column hierarchies independently. A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.

The `ShowAsRule` object has three properties:

- `calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).
- `baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied. Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.
- `baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type. Not all calculations require this field.

The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.
We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.
The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.

![A PivotTable showing the percentages of fruit sales relative to the grand total for both individual farms and individual fruit types within each farm.](../images/excel-pivots-showas-percentage.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {

        // Show the crates of each fruit type sold at the farm as a percentage of the column's total.
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Percentage of Total Farm Sales";
    });
});
```

The previous example set the calculation to the column, relative to the field of an individual row hierarchy. When the calculation relates to an individual item, use the `baseItem` property.

The following example shows the `differenceFrom` calculation. It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.
The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).

![A PivotTable showing the differences of fruit sales between “A Farms” and the others. This shows both the difference in total fruit sales of the farms and the sales of types of fruit. If “A Farms” did not sell a particular type of fruit, “#N/A” is displayed.](../images/excel-pivots-showas-differencefrom.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {
        // Show the difference between crate sales of the "A Farms" and the other farms.
        // This difference is both aggregated and shown for individual fruit types (where applicable).
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
        farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Difference from A Farms";
    });
});
```

## PivotTable layouts

A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data. You access the layout to determine the ranges where data is stored.

The following diagram shows which layout function calls correspond to which ranges of the PivotTable.

![A diagram showing which sections of a PivotTable are returned by the layout's get range functions.](../images/excel-pivots-layout-breakdown.png)

The following code demonstrates how to get the last row of the PivotTable data by going through the layout. Those values are then summed together for a grand total.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
        masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

PivotTables have three layout styles: Compact, Outline, and Tabular. We’ve seen the compact style in the previous examples.

The following examples use the outline and tabular styles, respectively. The code sample shows how to cycle between the different layouts.

### Outline layout

![A PivotTable using the outline layout.](../images/excel-pivots-outline-layout.png)

### Tabular layout

![A PivotTable using the tabular layout.](../images/excel-pivots-tabular-layout.png)

## Change hierarchy names

Hierarchy fields are editable. The following code demonstrates how to change the displayed names of two data hierarchies.

```js
Excel.run(function (context) {
    var dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    return context.sync().then(function () {
        // changing the displayed names of these entries
        dataHierarchies.items[0].name = "Farm Sales";
        dataHierarchies.items[1].name = "Wholesale";
    });
});
```

## Delete a PivotTable

PivotTables are deleted by using their name.

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## See also

- [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
- [Excel JavaScript API Reference](/javascript/api/excel)
