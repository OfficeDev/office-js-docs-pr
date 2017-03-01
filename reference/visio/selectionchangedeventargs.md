# SelectionChangedEventArgs Object (JavaScript API for Visio)

Applies to: _Visio Online_

Provides information about the shape collection that raised the SelectionChanged event.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|shapeNames|string[]|Gets the array of shape names that raised the SelectionChanged event.|
|pageName|string|Gets the name of the page which has the ShapeCollection object that raised the SelectionChanged event.|

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
	     	eventResult1 = document1.onSelectionChanged.add(
		function (args){
			       console.log("Selected Shape Name: "+args.shapeNames[0]);
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
