# Filter object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Excel for iOS, Office 2016_

Manages the filtering of a table's column.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|criteria|[FilterCriteria](filtercriteria.md)|The currently applied filter on the given column. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Apply the given filter criteria on the given column. The same functionality can be achieved with any of the following helper methods.|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Apply a "Bottom Item" filter to the column for the given number of elements.|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Apply a "Bottom Percent" filter to the column for the given percentage of elements.|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Apply a "Cell Color" filter to the column for the given color.|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|Apply a "Icon" filter to the column for the given criteria strings.|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|Apply a "Dynamic" filter to the column.|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Apply a "Font Color" filter to the column for the given color.|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Apply a "Icon" filter to the column for the given icon.|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Apply a "Top Item" filter to the column for the given number of elements.|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Apply a "Top Percent" filter to the column for the given percentage of elements.|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|Apply a "Values" filter to the column for the given values.|
|[clear()](#clear)|void|Clear the filter on the given column.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


### apply(criteria: FilterCriteria)
Apply the given filter criteria on the given column. The same functionality can be achieved with any of the following helper methods. 

#### Syntax
```js
filterObject.apply(criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|criteria|FilterCriteria|The criteria to apply.|

#### Returns
void

#### Example
The followng example demostrate how to apply a custom filter with the generic apply() method.

```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    var filterCriteria = { 
		filterOn: Excel.FilterOn.custom,
		criterion1: ">50",
		operator: Excel.FilterOperator.and,
		criterion2: "<100"
    	} 
    column.filter.apply(filterCriteria);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomItemsFilter(count: number)
Apply a "Bottom Item" filter to the column for the given number of elements.

#### Syntax
```js
filterObject.applyBottomItemsFilter(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|The number of elements from the bottom to show.|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomPercentFilter(percent: number)
Apply a "Bottom Percent" filter to the column for the given percentage of elements.

#### Syntax
```js
filterObject.applyBottomPercentFilter(percent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|percent|number|The percentage of elements from the bottom to show.|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyCellColorFilter(color: string)
Apply a "Cell Color" filter to the column for the given color.


#### Syntax
```js
filterObject.applyCellColorFilter(color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|color|string|The background color of the cells to show.|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCellColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)
Apply a "Icon" filter to the column for the given criteria strings.

#### Syntax
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|criteria1|string|The first criteria string.|
|criteria2|string|Optional. The second criteria string.|
|oper|FilterOperator|Optional. The operator that describes how the two criteria are joined.|

#### Returns
void


#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCustomFilter('>50','<100','and');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyDynamicFilter(criteria: string)
Apply a "Dynamic" filter to the column.

#### Syntax
```js
filterObject.applyDynamicFilter(criteria);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|criteria|string|The dynamic criteria to apply.  Possible values are: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyFontColorFilter(color: string)
Apply a "Font Color" filter to the column for the given color.

#### Syntax
```js
filterObject.applyFontColorFilter(color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|color|string|The font color of the cells to show.|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyFontColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyIconFilter(icon: Icon)
Apply a "Icon" filter to the column for the given icon.

#### Syntax
```js
filterObject.applyIconFilter(icon);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|icon|Icon|The icons of the cells to show.|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyIconFilter(Excel.icons.fiveArrows.yellowDownInclineArrow);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyTopItemsFilter(count: number)
Apply a "Top Item" filter to the column for the given number of elements.

#### Syntax
```js
filterObject.applyTopItemsFilter(count);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|count|number|The number of elements from the top to show.|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### applyTopPercentFilter(percent: number)
Apply a "Top Percent" filter to the column for the given percentage of elements.

#### Syntax
```js
filterObject.applyTopPercentFilter(percent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|percent|number|The percentage of elements from the top to show.|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyValuesFilter(values: ()[])
Apply a "Values" filter to the column for the given values.

#### Syntax
```js
filterObject.applyValuesFilter(values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|values|()[]|The list of values to show.|

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyValuesFilter(['a','b']);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
Clear the filter on the given column.

#### Syntax
```js
filterObject.clear();
```

#### Parameters
None

#### Returns
void

#### Example
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### load(param: object)
Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
