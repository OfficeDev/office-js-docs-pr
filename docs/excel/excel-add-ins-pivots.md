---
title: Work with Pivot Tables using the Excel JavaScript API
description: ''
ms.date: 7/11/2018
---



# Work with Pivot Tables using the Excel JavaScript API

Pivot tables streamline larger data sets. They allow the quick manipulation of grouped data. The Excel JavaScript API lets your add-in create pivot tables and interact with their components. 

If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user. 
See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools. 

This article provides code samples for common scenarios.
For the complete list of properties and methods the **PivotTable** and **PivotTableCollection** objects support, see [PivotTable API](https://dev.office.com/reference/add-ins/excel/pivottable) and [Pivot Table Collection API](https://dev.office.com/reference/add-ins/excel/pivottablecollection).

> [!NOTE]
> These samples use APIs currently available only in public preview (beta). To run these samples, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

## Hierarchies

Pivot tables are organized based on four hierarchy categories: row, column, data, and filter. The following data describing fruit sales from various farms will be used in throughout this article.

![A collection of fruit sales of different types from different farms.](../images/excel-pivots-raw-data.png)

This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**. Each hierarchy can only exist in one of the four categories. If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.

Row and column hierarchies define how data will be grouped. For example, having a row hierarchy of **Farms** groups together all data sets from the same farm. The choice between row and column hierarchy defines the orientation of the pivot table.

Data hierarchies are the values to be aggregated based on the row and column hierarchies. A pivot table with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.

Filter hierarchies include or exclude data from the pivot based on values within that filtered type. A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.

Here is the farm data again, alongside a pivot table. The pivot table is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected). 

![A selection of fruit sales data next to a pivot table with row, data, and filter hierarchies.](../images/excel-pivot-table-and-data.png)

This pivot table could be generated through the JavaScript API or through the Excel UI. Both options allow for further manipulation through add-ins.

## Create a pivot table

Pivot tables need a name, source, and destination. The source can be a `Range`, `string`, or `Table`. The destination can either be a `Range` or `string`. 
The following samples show various pivot table creation techniques. All three samples feature a pivot table named **Farm Sales** created on a worksheet called **PivotWorksheet** at cell **A2**. Its data comes from the worksheet **DataWorksheet** across the range **A1:E21**. 

### Create a pivot table with strings
```ts
await Excel.run(async (context) => {
	context.workbook.worksheets.getActiveWorksheet()
		.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

	await context.sync();
});
```

### Create a pivot table with Range objects
```ts
await Excel.run(async (context) => {	
	const myRange = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
	const myRange2 = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
	context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", myRange, myRange2);
	
	await context.sync();
});
```

### Create a pivot table at the workbook level
```ts
await Excel.run(async (context) => {
	context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

	await context.sync();
});
```

## Use an existing pivot table
Manually created pivot tables are also accessible through the pivot table collection of the workbook or of individual worksheets. 
The following code gets the first pivot table in the worksheet. It then gives the table a name for easy future reference.

```ts
await Excel.run(async (context) => {
	const pivotTableList = context.workbook.worksheets.getActiveWorksheet().pivotTables;
	pivotTableList.load("no-properties-needed");
	await context.sync();

	const pivotTable = pivotTableList.items[0];
	pivotTable.name = "My Pivot";
	await context.sync();
});
```

## Add rows and columns to a pivot table

Rows and columns pivot the data around those fields’ values.

Adding the **Farm** column pivots all the sales around each farm. Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.

![A pivot table with a Farm column and Type and Classification rows.](../images/excel-pivots-table-rows-and-columns.png)

```ts
await Excel.run(async (context) => {
	const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

	pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

	await context.sync();
});
```

You can also have a pivot table with only rows or columns.

```ts
await Excel.run(async (context) => {
	const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
	
	await context.sync();
});
```

## Add data hierarchies to the pivot table

Data hierarchies fill the pivot table with information to combine based on the rows and columns. 
Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column. 
In the example, both **Farm** and **Type** are rows, with the crate sales as the data. 

![A pivot table showing the total sales of different fruit based on the farm they came from.](../images/excel-pivots-data-hierarchy.png)

```ts
await Excel.run(async (context) => {
	const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

	pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
	pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

	await context.sync();
});
```

## Change aggregation function

Data hierarchies have their values aggregated into a sum by default. The `summarizeBy` parameter defines this behavior based on an `AggregrationFunction` type. 
The following code samples changes the aggregation to be averages of the data.

```ts
	await Excel.run(async (context) => {
        const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
        pivotTable.dataHierarchies.load("no-properties-needed");
        await context.sync();

        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        await context.sync();
    });
```

## Pivot table layouts

A pivot table layout defines the placement of hierarchies and their data. You access the layout to determine the ranges where data is stored. 
The following code demonstrates how to get the last row from a pivot table by going through the layout. Those values are then summed together for a grand total.

```ts
    await Excel.run(async (context) => {
        const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
        const range = pivotTable.layout.getDataBodyRange();
        const grandTotalRange = range.getLastRow();
        grandTotalRange.load();
        await context.sync();
		
		const masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
        masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
        await context.sync();
    });
```

Pivot tables have three layout styles: Compact, Outline, and Tabular. We’ve seen the compact style in the previous examples. 
The following examples use the outline and tabular styles, respectively. The code sample shows how to cycle between the different layouts.

### Outline layout
![A pivot table using the outline layout.](../images/excel-pivots-outline-layout.png)

### Tabular layout
![A pivot table usingthe tabular layout.](../images/excel-pivots-tabular-layout.png)

```ts
await Excel.run(async (context) => {
	const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
	pivotTable.layout.load("layoutType");
	await context.sync();
	if (pivotTable.layout.layoutType === "Compact") {
		pivotTable.layout.layoutType = "Outline";
	} else if (pivotTable.layout.layoutType === "Outline") {
		pivotTable.layout.layoutType = "Tabular";
	} else {
		pivotTable.layout.layoutType = "Compact";
	}
	
	await context.sync();
});
```

## Change hierarchy names

Hierarchy fields are editable. The following code demonstrates how to change the displayed names of two data hierarchies.

```ts
await Excel.run(async (context) => {
	const dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
		.pivotTables.getItem("Farm Sales").dataHierarchies;
	dataHierarchies.load("name");
	await context.sync();
	
	dataHierarchies.items[0].name = "Farm Sales";
	dataHierarchies.items[1].name = "Wholesale";
	await context.sync();
});
```

## Delete a pivot table

Pivot tables are deleted using their name.

```ts
await Excel.run(async (context) => {
	context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

	await context.sync();
});
```
