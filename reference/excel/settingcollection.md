# SettingCollection Object (JavaScript API for Excel)

Represents a collection of worksheet objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|A collection of setting objects. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(key: string, value: (any)[])](#addkey-string-value-any)|[Setting](setting.md)|Sets or adds the specified setting to the workbook.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of Settings in the collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|Gets a Setting entry via the key.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Setting](setting.md)|Gets a Setting entry via the key. If the Setting does not exist, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(key: string, value: (any)[])
Sets or adds the specified setting to the workbook.

#### Syntax
```js
settingCollectionObject.add(key, value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|The Key of the new setting.|
|value|(any)[]|The Value for the new setting.|

#### Returns
[Setting](setting.md)

### getCount()
Gets the number of Settings in the collection.

#### Syntax
```js
settingCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: string)
Gets a Setting entry via the key.

#### Syntax
```js
settingCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Key of the setting.|

#### Returns
[Setting](setting.md)

### getItemOrNullObject(key: string)
Gets a Setting entry via the key. If the Setting does not exist, will return a null object.

#### Syntax
```js
settingCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|The key of the setting.|

#### Returns
[Setting](setting.md)
