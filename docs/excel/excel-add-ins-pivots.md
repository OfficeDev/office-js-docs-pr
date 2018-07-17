---
title: Work with Pivot Tables using the Excel JavaScript API
description: ''
ms.date: 7/11/2018
---



# Work with Pivot Tables using the Excel JavaScript API

This article provides code samples that show how to perform common tasks with pivot tables using the Excel JavaScript API. 
For the complete list of properties and methods that the **PivotTable** and **PivotTableCollection** objects support, see [PivotTable Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/pivottable) and [Pivot Table Collection Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/pivottablecollection).

> [!NOTE]
> These samples use APIs currently available only in public preview (beta). To run these samples, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

## Create a pivot table


### Creating a pivot table with names
```ts
await Excel.run(async (context) => {
	context.workbook.worksheets.getActiveWorksheet()
		.pivotTables.add("Farm Sales", "Data!A1:E21", "Pivot!A2");

	await context.sync();
});
```

### Creating a pivot table with objects
```ts
await Excel.run(async (context) => {	
	var myRange = context.workbook.worksheets.getItem("Data").getRange("A1:E21");
	var myRange2 = context.workbook.worksheets.getItem("Pivot").getRange("A1");
	context.workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", myRange, myRange2);
	
	await context.sync();
});
```

### Creating a pivot table from a workbook
```ts
await Excel.run(async (context) => {
	context.workbook.pivotTables.add("Farm Sales", "Data!A1:E21", "Pivot!A2");

	await context.sync();
});
```

## Adding a column to a pivot table
```ts
await Excel.run(async (context) => {
	let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
	pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
	
	await context.sync();
});
```

## Adding rows to a pivot table
```ts
await Excel.run(async (context) => {
	let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
	pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
	
	await context.sync();
});
```

## Adding data hierarchies to the pivot table
```ts
await Excel.run(async (context) => {
	let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
	pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
	pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

	await context.sync();
});
```

## Rotating through the pivot table layouts
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

## Changing hierarchy names
```ts
await Excel.run(async (context) => {
	let pivotFields = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales").dataHierarchies
	pivotFields.load();
	await context.sync();
	
	pivotFields.items[0].name = "Farm Sales"
	pivotFields.items[1].name = "Wholesale"
	await context.sync();
});
```

## Deleting a pivot table
```ts
await Excel.run(async (context) => {
	context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

	await context.sync();
});
```
