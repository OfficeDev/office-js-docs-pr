# WorksheetProtectionOptions object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Excel for iOS, Office 2016_

Represents the options in sheet protection.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|allowAutoFilter|bool|Represents the worksheet protection option of allowing using auto filter feature.|
|allowDeleteColumns|bool|Represents the worksheet protection option of allowing deleting columns.|
|allowDeleteRows|bool|Represents the worksheet protection option of allowing deleting rows.|
|allowFormatCells|bool|Represents the worksheet protection option of allowing formatting cells.|
|allowFormatColumns|bool|Represents the worksheet protection option of allowing formatting columns.|
|allowFormatRows|bool|Represents the worksheet protection option of allowing formatting rows.|
|allowInsertColumns|bool|Represents the worksheet protection option of allowing inserting columns.|
|allowInsertHyperlinks|bool|Represents the worksheet protection option of allowing inserting hyperlinks.|
|allowInsertRows|bool|Represents the worksheet protection option of allowing inserting rows.|
|allowPivotTables|bool|Represents the worksheet protection option of allowing using pivot table feature.|
|allowSort|bool|Represents the worksheet protection option of allowing using sort feature.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


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

#### Examples
This example loads the protection options for the active worksheet.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection options: " + worksheet.protection.options);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
