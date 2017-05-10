# CustomXmlPart Object (JavaScript API for Excel)

Represents a custom XML part object in a workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|The custom XML part's ID. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|namespaceUri|string|The custom XML part's namespace URI. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the custom XML part.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getXml()](#getxml)|string|Gets the custom XML part's full XML content.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[setXml(xml: string)](#setxmlxml-string)|void|Sets the custom XML part's full XML content.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the custom XML part.

#### Syntax
```js
customXmlPartObject.delete();
```

#### Parameters
None

#### Returns
void

### getXml()
Gets the custom XML part's full XML content.

#### Syntax
```js
customXmlPartObject.getXml();
```

#### Parameters
None

#### Returns
string

### setXml(xml: string)
Sets the custom XML part's full XML content.

#### Syntax
```js
customXmlPartObject.setXml(xml);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|xml|string|XML content for the part.|

#### Returns
void
