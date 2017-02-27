# DataRefreshCompleteEventArgs Object (JavaScript API for Visio)

Applies to: _Visio Online_

Provides information about the document that raised the DataRefreshComplete event.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|success|bool|Gets the successfailure of the DataRefreshComplete event.|
|document|[Document](document.md)|Gets the document object that raised the DataRefreshComplete event.|

_See property access [examples.](#property-access-examples)_

## Relationships
None

## Methods
None

### Property access examples
```js
Visio.run(function (ctx) { 
  var document1= ctx.document;
               var page = document1.getActivePage();
	     eventResult1 = document1.onDataRefreshComplete.add(
	function (args){
	       console.log("Data Refresh Result: "+args.success);
	});

	return ctx.sync().then(function () {
	   console.log("Success");
	});
}).catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
	console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```
