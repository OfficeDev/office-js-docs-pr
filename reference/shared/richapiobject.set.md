

# Rich API object set method
Sets multiple properties of an object at once by passing either another object of the same type or a JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.

## Method Details

### set(properties: object, options: object)
The *non-read-only* properties of the object on which the method is called are set to the same values as the corresponding properties of the passed-in object.
If the `properties` parameter is a JavaScript object, then properties in the passed-in object that correspond to a read-only property in the object on which the method is called are either ignored or cause an exception, depending on the `options` parameter.

#### Syntax

```js
someRichAPIObject.set(properties[, options]);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|properties|object|Either an object *of the same Office type* on which the method is called, or a JavaScript object of property names and values that mirrors the structure of the properties of the object type on which the method is called.|
|options|object|Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)|

#### Returns

void    

#### Examples

The following example sets several format properties with a JavaScript object. 

```js
Excel.run(function (ctx) { 
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    })
    range.format.autofitColumns();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

The following example sets the properties of one range by copying the properties of another range. Note that the source object must be loaded first.

```js
Excel.run(function (ctx) { 
    var sheet = context.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");
    return ctx.sync(); 

    var targetRange = sheet.getRange("B7:E7");
    targetRange.set(sourceRange); 
    targetRange.format.autofitColumns();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
