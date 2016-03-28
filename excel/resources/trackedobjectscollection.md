# TrackedObjectsCollection object (JavaScript API for Office 2016)

Allows add-ins to manage range object references across sync() batches. Typically, Excel.run() allows you to maintain references across batches automatically, without having to track them explicitly. However, if an add-in scenario requires that a range object be tracked and adjusted manually to reflect the current state of the underlying Excel range, then this collection could be used to mark such objects for tracking. Note that if a range object is marked for tracking, it needs to be explicitly removed upon usage to free up the memory in Excel, even in an error case.

## Properties
None.

## Relationships

None

## Methods

The trackedObjectsCollection object has the following methods defined.

| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[add(rangeObject: Range)](#addrangeobject-range)| Null             |Creates a new reference on a range.|
|[remove(rangeObject: Range)](#removerangeobject-range)| Null             |Remove a reference on the range.  |
|[removeAll()](#removeall)| Null|Removes all references created by the add-in on the device.|


## API Specification 

### add(rangeObject: range)
Add a range object to the trackedObjectsCollection. Any underlying changes across batch requests will be tracked and any follow-up updates will be applied to the current state of the range object. 

#### Syntax
```js
trackedObjectsCollection.add(rangeObject);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| The range object that needs to be added to the trackedObjectCollection.

#### Returns
Null

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	return ctx.sync(); 
});
```


### remove(rangeObject: range)

Remove a reference object from the collection. This frees up memory and resources that are required to maintain the state of the tracked object. Note that if a range object is marked for tracking, it needs to be explicitly removed even in the case of an error.

#### Syntax
```js
trackedObjectsCollection.remove(rangeObject);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| The range object that needs to be removed from the trackedObjectCollection.

#### Returns
Null

#### Examples


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjectsCollection.remove(range); 
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
	ctx.trackedObjectsCollection.add(range);
	ctx.load(range);
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjectsCollection.removeAll(); 
	return ctx.sync(); 
});
```
