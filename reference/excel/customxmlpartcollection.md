# CustomXmlPartCollection Object (JavaScript API for Excel)

A collection of custom XML parts.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[CustomXmlPart[]](customxmlpart.md)|A collection of customXmlPart objects. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(xml: string)](#addxml-string)|[CustomXmlPart](customxmlpart.md)|Adds a new custom XML part to the workbook.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getByNamespace(namespaceUri: string)](#getbynamespacenamespaceuri-string)|[CustomXmlPartScopedCollection](customxmlpartscopedcollection.md)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of CustomXml parts in the collection.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[CustomXmlPart](customxmlpart.md)|Gets a custom XML part based on its ID.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(id: string)](#getitemornullobjectid-string)|[CustomXmlPart](customxmlpart.md)|Gets a custom XML part based on its ID.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(xml: string)
Adds a new custom XML part to the workbook.

#### Syntax
```js
customXmlPartCollectionObject.add(xml);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|xml|string|XML content. Must be a valid XML fragment.|

#### Returns
[CustomXmlPart](customxmlpart.md)

### getByNamespace(namespaceUri: string)
Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.

#### Syntax
```js
customXmlPartCollectionObject.getByNamespace(namespaceUri);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|namespaceUri|string||

#### Returns
[CustomXmlPartScopedCollection](customxmlpartscopedcollection.md)

### getCount()
Gets the number of CustomXml parts in the collection.

#### Syntax
```js
customXmlPartCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(id: string)
Gets a custom XML part based on its ID.

#### Syntax
```js
customXmlPartCollectionObject.getItem(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|ID of the object to be retrieved.|

#### Returns
[CustomXmlPart](customxmlpart.md)

### getItemOrNullObject(id: string)
Gets a custom XML part based on its ID.

#### Syntax
```js
customXmlPartCollectionObject.getItemOrNullObject(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|ID of the object to be retrieved.|

#### Returns
[CustomXmlPart](customxmlpart.md)
