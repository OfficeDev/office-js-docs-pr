# Tracked objects collection (JavaScript API for Office 2016)
Tracked cbjects collection allows add-ins to add and remove temporary references on range.

_Applies to: Office 2016_

## Properties
None.

## Relationships

None

## Methods

The TrackedObjects Collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(rangeObject: Range)](#addrangeobject-range)| Null             |Creates a new reference on a range.||
|[remove(rangeObject: Range)](#removerangeobject-range)| Null             |Remove a reference on the range.  ||
|[removeAll()](#removeall)| Null|Removes all references created by the add-in on the device.||


## API Specification 

### add(rangeObject: range)
Add a range object to the trackedObjectsCollection. 

#### Syntax
```js
trackedObjectsCollection.add(rangeObject);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| The Range Object which needs to be added to the trackedObjectCollection.

#### Returns
Null

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjects.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	return ctx.sync(); 
});
```


### remove(rangeObject: range)

Remove a reference object from the collection. 

#### Syntax
```js
trackedObjectsCollection.remove(rangeObject);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| The Range Object which needs to be removed from the trackedObjectCollection.

#### Returns
Null

#### Examples


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjects.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjects.remove(range); 
	return ctx.sync(); 
});
```

### removeAll(rangeObject: range)

Removes all references created by the add-in on the device.

#### Syntax
```js
trackedObjectsCollection.removeAll();
```

#### Parameters

None

#### Returns
Null

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2";
	var ctx = new Excel.RequestContext();
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	ctx.trackedObjects.add(range);
	ctx.load(range);
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjects.removeAll(); 
	return ctx.sync(); 
});
```

