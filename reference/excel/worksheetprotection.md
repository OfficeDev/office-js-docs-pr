# WorksheetProtection object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Excel for iOS, Office 2016_

Represents the protection of a sheet object.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|protected|bool|Indicates if the worksheet is protected. Read-Only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Sheet protection options. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object with protection details of the sheet.|
|[protect(options: WorksheetProtectionOptions, password: string)](#protectoptions-worksheetprotectionoptions-password-string)|void|Protect a worksheet. It throws if the worksheet has been protected.|
|[unprotect(password: string)](#unprotectpassword-string)|void|Unprotect a worksheet|

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
This example loads the protection details for the active worksheet.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection status: " + worksheet.protection.protected);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### protect(options: WorksheetProtectionOptions, password: string)
Protect a worksheet with optional protection policies and an optional password. It throws an exception if the worksheet has been protected. 

When options are specified, individual policies can be toggled enabled or diabled. If a policy isn't specified, then its enabled by default. 

#### Syntax
```js
worksheetProtectionObject.protect(options, password);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|options|WorksheetProtectionOptions|Optional. sheet protection options.|
|password|string|Optional. sheet protection password.|

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
### unprotect(password: string)
Unprotect a worksheet with an optional password which was used to protect the sheet. It throws an exception if the password supplied is wrong.

#### Syntax
```js
worksheetProtectionObject.unprotect(password);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|password|string|Optional. sheet protection password.|

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var sheet = ctx.workbook.worksheets.getItem("Sheet1");	
	sheet.protection.unprotect();
	return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```