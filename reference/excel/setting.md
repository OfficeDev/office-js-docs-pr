# Setting Object (JavaScript API for Excel)

Setting represents a key-value pair of a setting persisted to the document.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|key|string|Returns the key that represents the id of the Setting. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Represents the value stored for this setting.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the setting.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the setting.

#### Syntax
```js
settingObject.delete();
```

#### Parameters
None

#### Returns
void
