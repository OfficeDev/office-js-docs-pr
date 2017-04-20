# TableSort Object (JavaScript API for Excel)

Manages sorting operations on Table objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|matchCase|bool|Represents whether the casing impacted the last sort of the table. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|method|string|Represents Chinese character ordering method last used to sort the table. Read-only. Possible values are: PinYin, StrokeCount.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|fields|[SortField](sortfield.md)|Represents the current conditions used to last sort the table. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[apply(fields: SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|Perform a sort operation.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[reapply()](#reapply)|void|Reapplies the current sorting parameters to the table.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### apply(fields: SortField[], matchCase: bool, method: string)
Perform a sort operation.

#### Syntax
```js
tableSortObject.apply(fields, matchCase, method);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|fields|SortField[]|The list of conditions to sort on.|
|matchCase|bool|Optional. Whether to have the casing impact string ordering.|
|method|string|Optional. The ordering method used for Chinese characters.  Possible values are: PinYin, StrokeCount|

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### clear()
Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.

#### Syntax
```js
tableSortObject.clear();
```

#### Parameters
None

#### Returns
void

### reapply()
Reapplies the current sorting parameters to the table.

#### Syntax
```js
tableSortObject.reapply();
```

#### Parameters
None

#### Returns
void
