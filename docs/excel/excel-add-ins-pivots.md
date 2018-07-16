---
title: Work with Pivot Tables using the Excel JavaScript API
description: ''
ms.date: 7/11/2018
---



# Work with Pivot Tables using the Excel JavaScript API

This article provides code samples that show how to perform common tasks with pivot tables using the Excel JavaScript API. 
For the complete list of properties and methods that the **PivotTable** and **PivotTableCollection** objects support, see [PivotTable Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/pivottable) and [Chart Collection Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/pivottablecollection).

## Create a pivot table


### Creating the pivot table with names
```ts
await Excel.run(async (context) => {
	context.workbook.worksheets.getActiveWorksheet()
		.pivotTables.add("Farm Sales", "Data!A1:E21", "Pivot!A2");

	await context.sync();
});
```

### Creating the pivot table with objects
```ts
await Excel.run(async (context) => {	
	var myRange = context.workbook.worksheets.getItem("Data").getRange("A1:E21");
	var myRange2 = context.workbook.worksheets.getItem("Pivot").getRange("A1");
	context.workbook.worksheets.getItem("Pivot").pivotTables.add("Farm Sales", myRange, myRange2);
	
	await context.sync();
});
```

### Creating the pivot table from a workbook
```ts
await Excel.run(async (context) => {
	context.workbook.pivotTables.add("Farm Sales", "Data!A1:E21", "Pivot!A2");

	await context.sync();
});
```