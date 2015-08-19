# Reference Collection
Reference collection allows add-ins to add and remove temporary references on range.

## Properties
None.

## Relationships

None

## Methods

The Reference collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(rangeObject: Range)](#addrangeobject-range)| Null             |Creates a new reference on a range.  ||
|[remove(rangeObject: Range)](#removerangeobject-range)| Null             |Remove a reference on the range.  ||


## API Specification 

### add(rangeObject: range)
Add a range object to the reference collection. 

#### Syntax
```js
referenceCollection.add(rangeObject);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| The Range Object which needs to be added to the reference collection.

#### Returns
Null

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.references.add(range);
ctx.load(range);

ctx.executeAsync().then(function () {
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.executeAsync().then();
});
```
[Back](#methods)

### remove(rangeObject: range)

Remove a reference object from the collection. 

#### Syntax
```js
referenceCollection.remove(rangeObject);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| The Range Object which needs to be removed from the reference collection.

#### Returns
Null

#### Examples

```js
```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.references.add(range);
ctx.load(range);

ctx.executeAsync().then(function () {
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.references.remove(range); 
	ctx.executeAsync().then();
});
```
[Back](#methods)