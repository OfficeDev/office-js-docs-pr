# SettingCollection object (JavaScript API for Excel)

Represents a collection of worksheet objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|A collection of setting objects. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|Gets a Setting entry via the key.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: string)](#getitemornullkey-string)|[Setting](setting.md)|Gets a Setting entry via the key. If the Setting does not exist, the returned object's isNull property will be true.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[set(key: string, value: string)](#setkey-string-value-string)|[Setting](setting.md)|Sets or adds the specified setting to the workbook.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getItem(key: string)
Gets a Setting entry via the key.

#### Syntax
```js
settingCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Key of the setting.|

#### Returns
[Setting](setting.md)

### getItemOrNull(key: string)
Gets a Setting entry via the key. If the Setting does not exist, the returned object's isNull property will be true.

#### Syntax
```js
settingCollectionObject.getItemOrNull(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|The key of the setting.|

#### Returns
[Setting](setting.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### set(key: string, value: string)
Sets or adds the specified setting to the workbook.

#### Syntax
```js
settingCollectionObject.set(key, value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|The Key of the new setting.|
|value|string|The Value for the new setting.|

#### Returns
[Setting](setting.md)
