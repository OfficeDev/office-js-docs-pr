# ShapeMouseEnterEventArgs Object (JavaScript API for Visio)

Applies to: _Visio Online_

Provides information about the shape that raised the MouseEnter event.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|shapeName|string|Gets the name of the shape object that raised the MouseEnter event.|
|pageName|string|Gets the name of the page which has the shape object that raised the MouseEnter event.|

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
	eventResult2 = document1.onMouseEnter.add(
			function (args){			
		                 console.log(Date.now()+":OnMouseEnter Event"+JSON.stringify(args));
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