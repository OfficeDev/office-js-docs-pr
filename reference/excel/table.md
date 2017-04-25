# Table Object (JavaScript API for Excel)

Represents an Excel table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|highlightFirstColumn|bool|Indicates whether the first column contains special formatting.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|highlightLastColumn|bool|Indicates whether the last column contains special formatting.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|id|int|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the table.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedColumns|bool|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedRows|bool|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showFilterButton|bool|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showHeaders|bool|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTotals|bool|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See important note about performance related to table with [formulas](#setting-formulas)_



## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|columns|[TableColumnCollection](tablecolumncollection.md)|Represents a collection of all the columns in the table. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rows|[TableRowCollection](tablerowcollection.md)|Represents a collection of all the rows in the table. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[TableSort](tablesort.md)|Represents the sorting for the table. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current table. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|Clears all the filters currently applied on the table.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToRange()](#converttorange)|[Range](range.md)|Converts the table into a normal range of cells. All data is preserved.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Deletes the table.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the table.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with header row of the table.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire table.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with totals row of the table.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapplyFilters()](#reapplyfilters)|void|Reapplies all the filters currently on the table.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clearFilters()
Clears all the filters currently applied on the table.

#### Syntax
```js
tableObject.clearFilters();
```

#### Parameters
None

#### Returns
void

### convertToRange()
Converts the table into a normal range of cells. All data is preserved.

#### Syntax
```js
tableObject.convertToRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.convertToRange();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### delete()
Deletes the table.

#### Syntax
```js
tableObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getDataBodyRange()
Gets the range object associated with the data body of the table.

#### Syntax
```js
tableObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableDataRange = table.getDataBodyRange();
	tableDataRange.load('address')
	return ctx.sync().then(function() {
			console.log(tableDataRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getHeaderRowRange()
Gets the range object associated with header row of the table.

#### Syntax
```js
tableObject.getHeaderRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableHeaderRange = table.getHeaderRowRange();
	tableHeaderRange.load('address');
	return ctx.sync().then(function() {
		console.log(tableHeaderRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getRange()
Gets the range object associated with the entire table.

#### Syntax
```js
tableObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableRange = table.getRange();
	tableRange.load('address');	
	return ctx.sync().then(function() {
			console.log(tableRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getTotalRowRange()
Gets the range object associated with totals row of the table.

#### Syntax
```js
tableObject.getTotalRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableTotalsRange = table.getTotalRowRange();
	tableTotalsRange.load('address');	
	return ctx.sync().then(function() {
			console.log(tableTotalsRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### reapplyFilters()
Reapplies all the filters currently on the table.

#### Syntax
```js
tableObject.reapplyFilters();
```

#### Parameters
None

#### Returns
void
### Property access examples

Get a table by name. 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.load('index')
	return ctx.sync().then(function() {
			console.log(table.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Get a table by index.

```js
Excel.run(function (ctx) { 
	var index = 0;
	var table = ctx.workbook.tables.getItemAt(0);
	table.load('id')
	return ctx.sync().then(function() {
			console.log(table.id);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set table style. 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.name = 'Table1-Renamed';
	table.showTotals = false;
	table.style = 'TableStyleMedium2';
	table.load('tableStyle');
	return ctx.sync().then(function() {
			console.log(table.style);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Setting formulas

#### Common Pitfalls when Setting Formulas in Excel from Add-ins

by Zlatko Michailov  
Microsoft Corp.


This article points out three pitfalls that Excel add-in developers may encounter and ways to work around them. It is important to have an understanding of these scenarios especially because they don't cause the add-ins to fail under normal circumstances. The add-in may seem perfectly normal when used over a small range; however they may degrade in a linear fashion as the target range that add-in operates on grows over time.

The first two issues manifest when formulas are set on __table__ columns; in specifc, columns with formulas and columns with totals row.

##### Setting Formulas in Calculated Table Columns

This [article](https://support.office.com/en-us/article/Use-calculated-columns-in-an-Excel-table-873FBAC6-7110-4300-8F6F-AAFA2EA11CE8)
gives an overview of calculated columns.

The key feature is in Step 4:

> When you press Enter, the formula is automatically filled into all cells of the column — above as well as below the cell where you entered the formula.
> The formula is the same for each row, but since it's a structured reference, Excel knows internally which row is which.

That means that every single formula update may get multiplied N times where N is the number of rows in the table.

Users may not notice a substantial lag when dealing with a table with 1,000 rows, but interacting with a table containing 10,000 such rows may lead to  degraded experience.

Luckily, Excel's automatic column calculation is clever enough, and you may not notice the above problem. For a column to get automatically recalculated, it has to either be empty or be entirely auto-calculated. If you break the "purity" of the column by inserting a value (not a formula) in any cell, Excel will not try to auto-recalculate it. Also, if you are trying to set the formula that Excel has already set in that column, the recalculation would be a no-op.

Example, let's say you are setting the formula `=B2+C2` on cell `A2`. If the column is empty, Excel will calculate all the cells of this column _adjusting the row index_. Then, when you move to the next row, and you set the formula `=B3+C3` on `A3`, there will be no column recalculation, because this formula is already auto-set on the whole column.

However, if you want your column to represent a function of the row index, e.g. `=i * i` where _i_ is the row index,
not only will this cause a whole column recalculation on every update, but you will also end up with a column that shows the same (last) formula.

##### Setting Formulas on a Table with a "Totals" Row

Setting formulas on tables with totals row enabled may sometimes cause performance issues. It is important to mention that even a default Totals row, i.e. one that has a static value in the left-most cell and a `Count` on the right-most cells and having all cells in between `None`, could repro the problem. 

While there is a simpler workaround - set all the formulas, and then add the totals row on the table, the generic workaround pattern that is recommended for both of above issues is to use a plain range while setting formulas, and then to convert that range into a table.

Here is a generic function that updates a range of data and creates a table on the target range. 

```js
function createAndPopulateTable(context, worksheetName, rangeAddress, hasHeaderRow, headerValues, bodyFormulas, tableCustomizer) {
    var worksheet = context.workbook.worksheets.getItem(worksheetName);

    // Calculate table-, body-, and header- ranges
    var tableRange = worksheet.getRange(rangeAddress);
    var bodyRange = tableRange;
    if (hasHeaderRow) {
        bodyRange = tableRange.getResizedRange(-1, 0).getOffsetRange(1, 0);
        if (headerValues) {
            // Set header values
            var headerRange = tableRange.getRow(0);
            headerRange.values = headerValues;
        }
    }
    
    // Set body formulas
    bodyRange.formulas = bodyFormulas;

    return context.sync()
        .then(function() {
            // Create the table
            var table = context.workbook.tables.add(tableRange, hasHeaderRow);

            // Invoke the caller's customizer
            if (tableCustomizer) {
                tableCustomizer(table);
            }

            return context.sync();
        });
}
```

The above function is available online at [public location](https://gist.github.com/zlatko-michailov/2b0418c986d9da6ee0bdf7aa346d3a4f).

It can be used like this:
```js
    return Excel.run(function(context) {
        return createAndPopulateTable(context, "Sheet1", "B3:E6", true, [['Alpha', 'Beta', 'Gamma', 'Delta']], 
                    [ ['=1+1', null, null, '=B4'], 
                      ['=2+2', null, null, '=B5'],
                      ['=3+3', null, null, '=B6'] ],
                    function (table) {
                        table.style = 'TableStyleLight1';
                        table.showTotals = true;
                    });
    });
```

Automatic column calculation can be disabled in Excel desktop client (it is ON by default), however it is always ON in Excel Online.
Therefore, as an add-in developer, you should assume that it is ON for the majority of your add-in's users.


##### Getting a Range Object

This issue is specific to the JavaScript API implementation.

In order to correctly track the range during insertions and deletions of rows/columns, a binding is internally created every time a `Range` object is requested.
Later, when a cell is updated, all relevant bindings have to get notified to update themselves.

Thus, the following code (line 8), that seems benign from a general programming perspective escalates complexity quadraticly:
```js
    Excel.run(function(context) {
        var n = 10000;
        var worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load();

        var arr = [];
        for (var i = 2; i <= n + 1; i++) {
            var range = worksheet.getRange("C3:C" + (n + 1)); /* <-- PROBLEM! */
            arr.push(["=A" + i + " + B" + i]);
        }
        range.formulas = arr; 
        return context.sync();
    });
```

The workaround is to avoid the unnecessary get's to same `Range` object by taking the relevant line outside of the loop:
```js
    Excel.run(function(context) {
        var n = 10000;
        var worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load();

        var arr = [];
        var range = worksheet.getRange("C3:C" + (n + 1)); /* <-- OK */
        for (var i = 2; i <= n + 1; i++) {
            arr.push(["=A" + i + " + B" + i]);
        }
        range.formulas = arr; 
        return context.sync();
    });
```
