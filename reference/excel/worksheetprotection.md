# WorksheetProtection Object (JavaScript API for Excel)

Represents the protection of a sheet object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|protected|bool|Indicates if the worksheet is protected. Read-Only. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Sheet protection options. Read-Only. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[protect(options: WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|void|Protects a worksheet. Fails if the worksheet has been protected.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[unprotect()](#unprotect)|void|Unprotects a worksheet.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### protect(options: WorksheetProtectionOptions)
Protects a worksheet. Fails if the worksheet has been protected.

#### Syntax
```js
worksheetProtectionObject.protect(options);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|options|WorksheetProtectionOptions|Optional. sheet protection options.|

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var sheet = ctx.workbook.worksheets.getItem("Sheet1");
	var range = sheet.getRange("A1:B3").format.protection.locked = false;
	sheet.protection.protect({allowInsertRows:true});
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});

```
### unprotect()
Unprotects a worksheet.

#### Syntax
```js
worksheetProtectionObject.unprotect();
```

#### Parameters
None

#### Returns
void
