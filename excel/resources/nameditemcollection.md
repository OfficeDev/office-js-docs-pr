# NamedItemCollection

A collection of all the nameditem objects that are part of the workbook.

## [Properties](#getter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[NamedItem[]](nameditem.md)|A collection of namedItem objects. Read-only.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

### getItem(name: string)
Gets a nameditem object using its name

#### Syntax
```js
namedItemCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|nameditem name.|

#### Returns
[NamedItem](nameditem.md)

#### Examples

```js
var ctx = new Excel.RequestContext();
var nameditem = ctx.workbook.names.getItem(wSheetName);
nameditem.load(type);
ctx.executeAsync().then(function () {
		Console.log(nameditem.type);
});
```

```js
var ctx = new Excel.RequestContext();
var nameditem = ctx.workbook.names.getItemAt(0);
nameditem.load(name);
ctx.executeAsync().then(function () {
		Console.log(nameditem.name);
});
```
[Back](#methods)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples
```js

```

[Back](#methods)

### Getter Examples

```js
var ctx = new Excel.RequestContext();
var nameditems = ctx.workbook.names;
nameditems.load(items);
ctx.executeAsync().then(function () {
	for (var i = 0; i < nameditems.items.length; i++)
	{
		Console.log(nameditems.items[i].name);
		Console.log(nameditems.items[i].index);
	}
});
```

Get the number of nameditems.

```js
var ctx = new Excel.RequestContext();
var nameditems = ctx.workbook.names;
nameditems.load(count);
ctx.executeAsync().then(function () {
	Console.log("nameditems: Count= " + nameditems.count);
});

```


[Back](#properties)
