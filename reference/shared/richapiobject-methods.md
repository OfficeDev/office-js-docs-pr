# Load or set properties of an object

You can use the **object.load** method and the **object.set** method to load or set properties of an object in host-specific Office JavaScript APIs, such as the Excel JavaScript APIs and the Word JavaScript APIs.

## object.load method

> **Note**: This method is implemented only on objects in host-specific Office JavaScript APIs, such as the Excel JavaScript APIs and the Word JavaScript APIs. It is not implmented on objects in the Common APIs. For more information about the distinction between host-specific APIs and Common APIs, see [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office).

### Method Details

#### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide a host-relative loadOption object. Excel: [loadOption](/reference/excel/loadoption.md) Word: [loadOption](/reference/word/loadoption.md).|

#### Returns
void

#### Examples

The following example sets the properties of one Excel range by copying the properties of another range. Note that the source object must be loaded first. The example assumes there is data two ranges, B2:E2 and B7:E7, and that they are initially formatted differently.

```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange); 
            targetRange.format.autofitColumns();

            return ctx.sync()        
        })     
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

The following example sets the properties of one Word paragraph by copying the properties of another paragraph. Note that the source object must be loaded first. The example assumes there are at least two paragraphs, and that they are initially formatted differently.

```js
async function copyPropertiesFromParagraph() {
    await Word.run(async (context) => {
        const firstParagraph = context.document.body.paragraphs.getFirst();
        const secondParagraph = firstParagraph.getNext();
        firstParagraph.load("text, font/color, font/bold, leftIndent");

        await context.sync();

        secondParagraph.set(firstParagraph);

        await context.sync();
    });
}
```



## object.set method
Sets multiple properties of an object at once by passing either another object of the same Office type or a JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.

> **Note**: This method is implemented only on objects in host-specific Office JavaScript APIs, such as the Excel JavaScript APIs and the Word JavaScript APIs. It is not implmented on objects in the Common APIs. For more information about the distinction between host-specific APIs and Common APIs, see [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office).

### Method Details

#### set(properties: object, options: object)
The *non-read-only* properties of the object on which the method is called are set to the same values as the corresponding properties of the passed-in object.
If the `properties` parameter is a JavaScript object, then properties in the passed-in object that correspond to a read-only property in the object on which the method is called are either ignored or cause an exception, depending on the `options` parameter.

#### Syntax

```js
object.set(properties[, options]);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|properties|object|Either an object *of the same Office type* on which the method is called, or a JavaScript object of property names and values that mirrors the structure of the properties of the object type on which the method is called.|
|options|object|Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)|

#### Returns

void    

#### Examples

The following example sets several Excel format properties with a JavaScript object. The example assumes that there is data in range B2:E2.

```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sample");
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

The following example sets the properties of one Excel range by copying the properties of another range. Note that the source object must be loaded first. The example assumes there is data two ranges, B2:E2 and B7:E7, and that they are initially formatted differently.

```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange); 
            targetRange.format.autofitColumns();

            return ctx.sync()        
        })     
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

The following example sets several Word paragraph format properties with a JavaScript object. The example assumes that there is at least one paragraph.

```js
async function setMultiplePropertiesWithObject() {
    await Word.run(async (context) => {
        const paragraph = context.document.body.paragraphs.getFirst();
        paragraph.set({
            leftIndent: 30,
            font: {
                bold: true,
                color: 'red'
            }
        });

        await context.sync();
    });
}
```

The following example sets the properties of one Word paragraph by copying the properties of another paragraph. Note that the source object must be loaded first. The example assumes there are at least two paragraphs, and that they are initially formatted differently.

```js
async function copyPropertiesFromParagraph() {
    await Word.run(async (context) => {
        const firstParagraph = context.document.body.paragraphs.getFirst();
        const secondParagraph = firstParagraph.getNext();
        firstParagraph.load("text, font/color, font/bold, leftIndent");

        await context.sync();

        secondParagraph.set(firstParagraph);

        await context.sync();
    });
}
```