# ConditionalFormatCollection Object (JavaScript API for Excel)

Represents a collection of all the conditional formats that are overlap the range.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[ConditionalFormat[]](conditionalformat.md)|A collection of conditionalFormat objects. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(type: string)](#addtype-string)|[ConditionalFormat](conditionalformat.md)|Adds a new conditional format to the collection at the firsttop priority.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|[clearAll()](#clearall)|void|Clears all conditional formats active on the current specified range.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of conditional formats in the workbook. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ConditionalFormat](conditionalformat.md)|Returns a conditional format at the given index.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(type: string)
Adds a new conditional format to the collection at the firsttop priority.

#### Syntax
```js
conditionalFormatCollectionObject.add(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|string|The type of conditional format being added.  Possible values are: Custom, DataBar, ColorScale, IconSet|

#### Returns
[ConditionalFormat](conditionalformat.md)

#### Examples
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    conditionalFormat.iconOrNull.style = "YellowThreeArrows";
    return ctx.sync().then(function () {
        console.log("Added new yellow three arrow icon set.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```


### clearAll()
Clears all conditional formats active on the current specified range.

#### Syntax
```js
conditionalFormatCollectionObject.clearAll();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormats = range.conditionalFormats;
    var conditionalFormat = conditionalFormats.clearAll();
    return ctx.sync().then(function () {
        console.log("Cleared all conditional formats from this range.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```


### getCount()
Returns the number of conditional formats in the workbook. Read-only.

#### Syntax
```js
conditionalFormatCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

#### Examples
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    conditionalFormat.iconOrNull.style = Excel.IconSet.fourTrafficLights;
    var cfCount = range.conditionalFormats.getCount(); 

    return ctx.sync().then(function () {
        console.log("Count: " + cfCount.value);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### getItemAt(index: number)
Returns a conditional format at the given index.

#### Syntax
```js
conditionalFormatCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the conditional formats to be retrieved.|

#### Returns
[ConditionalFormat](conditionalformat.md)

#### Examples
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormats = range.conditionalFormats;
    var conditionalFormat = conditionalFormats.getItemAt(3);
    return ctx.sync().then(function () {
        console.log("Conditional Format at Item 3 Loaded");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

