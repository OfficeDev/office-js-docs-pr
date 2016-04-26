# Application Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

The Application object.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[createDocument(base64File: string)](#createdocumentbase64file-string)|[Document](document.md)|Creates a new document by using a base64 encoded .docx file.|WordApiWordApiDesktop, 1.3|

## Method Details


### createDocument(base64File: string)
Creates a new document by using a base64 encoded .docx file.

#### Syntax
```js
applicationObject.createDocument(base64File);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64File|string|Optional. Optional. The base64 encoded .docx file. The default value is null.|

#### Returns
[Document](document.md)
### Property access examples
```js
Excel.run(function (ctx) { 
	var application = ctx.workbook.application;
	application.load('calculationMode');
	return ctx.sync().then(function() {
		console.log(application.calculationMode);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

