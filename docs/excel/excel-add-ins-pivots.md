---
title: Work with Pivot Tables using the Excel JavaScript API
description: ''
ms.date: 7/11/2018
---



# Work with Pivot Tables using the Excel JavaScript API

Pivot tables streamline larger data sets. They allow the quick manipulation of grouped data. The Excel JavaScript API lets your add-in create pivot tables and interact with their components. This article provides code samples for common scenarios.
For the complete list of properties and methods the **PivotTable** and **PivotTableCollection** objects support, see [PivotTable API](https://dev.office.com/reference/add-ins/excel/pivottable) and [Pivot Table Collection API](https://dev.office.com/reference/add-ins/excel/pivottablecollection).

> [!NOTE]
> These samples use APIs currently available only in public preview (beta). To run these samples, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.


## Create a pivot table

Pivot tables need a name, source, and destination. The source can be a `Range`, `string`, or `Table`. The destination can either be a `Range` or `string`. 
The following samples show various pivot table creation techniques. All three samples feature a pivot table named **Farm Sales** created on a worksheet called **Pivot** at cell **A2**. Its data comes from the worksheet **Data** across the range **A1:E21**. 

### Create a pivot table with names
```ts
await Excel.run(async (context) => {
	context.workbook.worksheets.getActiveWorksheet()
		.pivotTables.add("Farm Sales", "Data!A1:E21", "Pivot!A2");

	await context.sync();
});
```

### Create a pivot table with objects
```ts
await Excel.run(async (context) => {	
	var myRange = context.workbook.worksheets.getItem("Data").getRange("A1:E21");
	var myRange2 = context.workbook.worksheets.getItem("Pivot").getRange("A2);
	context.workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", myRange, myRange2);
	
	await context.sync();
});
```

### Create a pivot table from a workbook
```ts
await Excel.run(async (context) => {
	context.workbook.pivotTables.add("Farm Sales", "Data!A1:E21", "Pivot!A2");

	await context.sync();
});
```

## Add rows and columns to a pivot table

Rows and columns pivot the data around those fields’ values. The following data describes fruit sales from various farms.

![A collection of fruit sales of different types from different farms](../images/excel-pivots-raw-data.png)

Adding the **Farm** column pivots all the sales around each farm. Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.

![A pivot table with a Farm column and Type and Classification rows.](../images/excel-pivots-table-rows-and-columns.png)

```ts
await Excel.run(async (context) => {
	let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

	pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

	await context.sync();
});
```

You can also have a pivot table with only rows or columns.

```ts
await Excel.run(async (context) => {
	let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
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
	let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

	pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
	pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

	await context.sync();
});
```

## Pivot table layouts

Pivot tables have three layout styles: Compact, Outline, and Tabular. We’ve seen the compact style in the previous examples. 
The following examples use the outline and tabular styles, respectively. The code sample shows how to cycle between the different layouts.

### Outline layout
![A pivot table using the outline layout.](../images/excel-pivots-outline-layout.png)

### Tabular layout
![A pivot table usingthe tabular layout.](../images/excel-pivots-tabular-layout.png)

```ts
await Excel.run(async (context) => {
	let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
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
	let pivotFields = context.workbook.worksheets.getActiveWorksheet()
		.pivotTables.getItem("Farm Sales").dataHierarchies;
	pivotFields.load();
	await context.sync();
	
	pivotFields.items[0].name = "Farm Sales";
	pivotFields.items[1].name = "Wholesale";
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
