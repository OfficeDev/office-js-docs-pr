# RangeReference object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Excel for iOS, Office 2016_

Represents a string reference of the form SheetName!A1:B5, or a global or local named range

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|address|string|The worksheet containing the current range.|

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
