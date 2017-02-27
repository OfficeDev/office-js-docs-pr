# PageLoadCompleteEventArgs Object (JavaScript API for Visio)

Applies to: _Visio Online_

Provides information about the page that raised the PageLoadComplete event.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|pageName|string|Gets the name of the page that raised the PageLoad event.|
|success|bool|Gets the success or failure of the PageLoadComplete event.|

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
	     	eventResult1 = document1.onPageLoadComplete.add(
			function (args){
			       console.log("Page name: "+args.pageName);
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
